VERSION 5.00
Begin VB.Form PARAVEN2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DEFINICION FORMULARIO DE VENTANA"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0FFD0&
      BorderStyle     =   0  'None
      Caption         =   "Vencimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   12000
      Begin VB.Frame Frame14 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         Height          =   5420
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   11685
         Begin VB.Frame Frame15 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame15"
            Height          =   19200
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   11475
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   65
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   1056
               Top             =   19500
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   65
                  Left            =   320
                  TabIndex        =   1070
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   65
                  Left            =   4050
                  TabIndex        =   1069
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   65
                  Left            =   4250
                  TabIndex        =   1068
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   65
                  Left            =   4460
                  TabIndex        =   1067
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   65
                  Left            =   8260
                  TabIndex        =   1066
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   65
                  Left            =   8660
                  TabIndex        =   1065
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   65
                  Left            =   9050
                  TabIndex        =   1064
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   65
                  Left            =   9250
                  TabIndex        =   1063
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   65
                  Left            =   9460
                  TabIndex        =   1062
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   65
                  Left            =   9660
                  TabIndex        =   1061
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   65
                  Left            =   11085
                  TabIndex        =   1060
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   65
                  Left            =   4850
                  TabIndex        =   1059
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   65
                  Left            =   5240
                  TabIndex        =   1058
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   65
                  Left            =   10870
                  TabIndex        =   1057
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   65
                  Left            =   0
                  TabIndex        =   1071
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   64
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   1040
               Top             =   19200
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   64
                  Left            =   5240
                  TabIndex        =   1054
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   64
                  Left            =   4850
                  TabIndex        =   1053
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   64
                  Left            =   11085
                  TabIndex        =   1052
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   64
                  Left            =   9660
                  TabIndex        =   1051
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   64
                  Left            =   9460
                  TabIndex        =   1050
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   64
                  Left            =   9250
                  TabIndex        =   1049
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   64
                  Left            =   9050
                  TabIndex        =   1048
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   64
                  Left            =   8660
                  TabIndex        =   1047
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   64
                  Left            =   8260
                  TabIndex        =   1046
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   64
                  Left            =   4460
                  TabIndex        =   1045
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   64
                  Left            =   4250
                  TabIndex        =   1044
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   64
                  Left            =   4050
                  TabIndex        =   1043
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   64
                  Left            =   320
                  TabIndex        =   1042
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   64
                  Left            =   10870
                  TabIndex        =   1041
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   64
                  Left            =   0
                  TabIndex        =   1055
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   63
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   1024
               Top             =   18900
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   63
                  Left            =   10870
                  TabIndex        =   1038
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   63
                  Left            =   5240
                  TabIndex        =   1037
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   63
                  Left            =   4850
                  TabIndex        =   1036
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   63
                  Left            =   11085
                  TabIndex        =   1035
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   63
                  Left            =   9660
                  TabIndex        =   1034
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   63
                  Left            =   9460
                  TabIndex        =   1033
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   63
                  Left            =   9250
                  TabIndex        =   1032
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   63
                  Left            =   9050
                  TabIndex        =   1031
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   63
                  Left            =   8660
                  TabIndex        =   1030
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   63
                  Left            =   8260
                  TabIndex        =   1029
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   63
                  Left            =   4460
                  TabIndex        =   1028
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   63
                  Left            =   4250
                  TabIndex        =   1027
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   63
                  Left            =   4050
                  TabIndex        =   1026
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   63
                  Left            =   320
                  TabIndex        =   1025
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   63
                  Left            =   0
                  TabIndex        =   1039
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   62
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   1008
               Top             =   18600
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   62
                  Left            =   10870
                  TabIndex        =   1022
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   62
                  Left            =   320
                  TabIndex        =   1021
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   62
                  Left            =   4050
                  TabIndex        =   1020
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   62
                  Left            =   4250
                  TabIndex        =   1019
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   62
                  Left            =   4460
                  TabIndex        =   1018
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   62
                  Left            =   8260
                  TabIndex        =   1017
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   62
                  Left            =   8660
                  TabIndex        =   1016
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   62
                  Left            =   9050
                  TabIndex        =   1015
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   62
                  Left            =   9250
                  TabIndex        =   1014
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   62
                  Left            =   9460
                  TabIndex        =   1013
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   62
                  Left            =   9660
                  TabIndex        =   1012
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   62
                  Left            =   11085
                  TabIndex        =   1011
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   62
                  Left            =   4850
                  TabIndex        =   1010
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   62
                  Left            =   5240
                  TabIndex        =   1009
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   62
                  Left            =   0
                  TabIndex        =   1023
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   61
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   992
               Top             =   18300
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   61
                  Left            =   10870
                  TabIndex        =   1006
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   61
                  Left            =   5240
                  TabIndex        =   1005
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   61
                  Left            =   4850
                  TabIndex        =   1004
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   61
                  Left            =   11085
                  TabIndex        =   1003
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   61
                  Left            =   9660
                  TabIndex        =   1002
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   61
                  Left            =   9460
                  TabIndex        =   1001
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   61
                  Left            =   9250
                  TabIndex        =   1000
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   61
                  Left            =   9050
                  TabIndex        =   999
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   61
                  Left            =   8660
                  TabIndex        =   998
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   61
                  Left            =   8260
                  TabIndex        =   997
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   61
                  Left            =   4460
                  TabIndex        =   996
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   61
                  Left            =   4250
                  TabIndex        =   995
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   61
                  Left            =   4050
                  TabIndex        =   994
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   61
                  Left            =   320
                  TabIndex        =   993
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   61
                  Left            =   0
                  TabIndex        =   1007
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   60
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   976
               Top             =   18000
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   60
                  Left            =   10870
                  TabIndex        =   990
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   60
                  Left            =   320
                  TabIndex        =   989
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   60
                  Left            =   4050
                  TabIndex        =   988
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   60
                  Left            =   4250
                  TabIndex        =   987
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   60
                  Left            =   4460
                  TabIndex        =   986
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   60
                  Left            =   8260
                  TabIndex        =   985
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   60
                  Left            =   8660
                  TabIndex        =   984
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   60
                  Left            =   9050
                  TabIndex        =   983
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   60
                  Left            =   9250
                  TabIndex        =   982
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   60
                  Left            =   9460
                  TabIndex        =   981
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   60
                  Left            =   9660
                  TabIndex        =   980
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   60
                  Left            =   11085
                  TabIndex        =   979
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   60
                  Left            =   4850
                  TabIndex        =   978
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   60
                  Left            =   5240
                  TabIndex        =   977
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   60
                  Left            =   0
                  TabIndex        =   991
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   59
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   960
               Top             =   17700
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   59
                  Left            =   320
                  TabIndex        =   974
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   59
                  Left            =   4050
                  TabIndex        =   973
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   59
                  Left            =   4250
                  TabIndex        =   972
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   59
                  Left            =   4460
                  TabIndex        =   971
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   59
                  Left            =   8260
                  TabIndex        =   970
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   59
                  Left            =   8660
                  TabIndex        =   969
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   59
                  Left            =   9050
                  TabIndex        =   968
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   59
                  Left            =   9250
                  TabIndex        =   967
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   59
                  Left            =   9460
                  TabIndex        =   966
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   59
                  Left            =   9660
                  TabIndex        =   965
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   59
                  Left            =   11085
                  TabIndex        =   964
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   59
                  Left            =   4850
                  TabIndex        =   963
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   59
                  Left            =   5240
                  TabIndex        =   962
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   59
                  Left            =   10870
                  TabIndex        =   961
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   59
                  Left            =   0
                  TabIndex        =   975
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   58
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   944
               Top             =   17400
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   58
                  Left            =   5240
                  TabIndex        =   958
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   58
                  Left            =   4850
                  TabIndex        =   957
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   58
                  Left            =   11085
                  TabIndex        =   956
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   58
                  Left            =   9660
                  TabIndex        =   955
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   58
                  Left            =   9460
                  TabIndex        =   954
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   58
                  Left            =   9250
                  TabIndex        =   953
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   58
                  Left            =   9050
                  TabIndex        =   952
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   58
                  Left            =   8660
                  TabIndex        =   951
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   58
                  Left            =   8260
                  TabIndex        =   950
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   58
                  Left            =   4460
                  TabIndex        =   949
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   58
                  Left            =   4250
                  TabIndex        =   948
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   58
                  Left            =   4050
                  TabIndex        =   947
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   58
                  Left            =   320
                  TabIndex        =   946
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   58
                  Left            =   10870
                  TabIndex        =   945
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   58
                  Left            =   0
                  TabIndex        =   959
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   57
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   928
               Top             =   17100
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   57
                  Left            =   10870
                  TabIndex        =   942
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   57
                  Left            =   5240
                  TabIndex        =   941
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   57
                  Left            =   4850
                  TabIndex        =   940
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   57
                  Left            =   11085
                  TabIndex        =   939
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   57
                  Left            =   9660
                  TabIndex        =   938
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   57
                  Left            =   9460
                  TabIndex        =   937
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   57
                  Left            =   9250
                  TabIndex        =   936
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   57
                  Left            =   9050
                  TabIndex        =   935
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   57
                  Left            =   8660
                  TabIndex        =   934
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   57
                  Left            =   8260
                  TabIndex        =   933
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   57
                  Left            =   4460
                  TabIndex        =   932
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   57
                  Left            =   4250
                  TabIndex        =   931
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   57
                  Left            =   4050
                  TabIndex        =   930
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   57
                  Left            =   320
                  TabIndex        =   929
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   57
                  Left            =   0
                  TabIndex        =   943
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   56
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   912
               Top             =   16800
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   56
                  Left            =   10870
                  TabIndex        =   926
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   56
                  Left            =   320
                  TabIndex        =   925
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   56
                  Left            =   4050
                  TabIndex        =   924
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   56
                  Left            =   4250
                  TabIndex        =   923
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   56
                  Left            =   4460
                  TabIndex        =   922
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   56
                  Left            =   8260
                  TabIndex        =   921
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   56
                  Left            =   8660
                  TabIndex        =   920
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   56
                  Left            =   9050
                  TabIndex        =   919
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   56
                  Left            =   9250
                  TabIndex        =   918
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   56
                  Left            =   9460
                  TabIndex        =   917
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   56
                  Left            =   9660
                  TabIndex        =   916
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   56
                  Left            =   11085
                  TabIndex        =   915
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   56
                  Left            =   4850
                  TabIndex        =   914
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   56
                  Left            =   5240
                  TabIndex        =   913
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   56
                  Left            =   0
                  TabIndex        =   927
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   55
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   896
               Top             =   16500
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   55
                  Left            =   320
                  TabIndex        =   910
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   55
                  Left            =   4050
                  TabIndex        =   909
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   55
                  Left            =   4250
                  TabIndex        =   908
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   55
                  Left            =   4460
                  TabIndex        =   907
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   55
                  Left            =   8260
                  TabIndex        =   906
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   55
                  Left            =   8660
                  TabIndex        =   905
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   55
                  Left            =   9050
                  TabIndex        =   904
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   55
                  Left            =   9250
                  TabIndex        =   903
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   55
                  Left            =   9460
                  TabIndex        =   902
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   55
                  Left            =   9660
                  TabIndex        =   901
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   55
                  Left            =   11085
                  TabIndex        =   900
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   55
                  Left            =   4850
                  TabIndex        =   899
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   55
                  Left            =   5240
                  TabIndex        =   898
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   55
                  Left            =   10870
                  TabIndex        =   897
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   55
                  Left            =   0
                  TabIndex        =   911
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   54
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   880
               Top             =   16200
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   54
                  Left            =   5240
                  TabIndex        =   894
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   54
                  Left            =   4850
                  TabIndex        =   893
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   54
                  Left            =   11085
                  TabIndex        =   892
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   54
                  Left            =   9660
                  TabIndex        =   891
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   54
                  Left            =   9460
                  TabIndex        =   890
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   54
                  Left            =   9250
                  TabIndex        =   889
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   54
                  Left            =   9050
                  TabIndex        =   888
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   54
                  Left            =   8660
                  TabIndex        =   887
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   54
                  Left            =   8260
                  TabIndex        =   886
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   54
                  Left            =   4460
                  TabIndex        =   885
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   54
                  Left            =   4250
                  TabIndex        =   884
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   54
                  Left            =   4050
                  TabIndex        =   883
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   54
                  Left            =   320
                  TabIndex        =   882
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   54
                  Left            =   10870
                  TabIndex        =   881
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   54
                  Left            =   0
                  TabIndex        =   895
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   53
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   864
               Top             =   15900
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   53
                  Left            =   320
                  TabIndex        =   878
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   53
                  Left            =   4050
                  TabIndex        =   877
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   53
                  Left            =   4250
                  TabIndex        =   876
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   53
                  Left            =   4460
                  TabIndex        =   875
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   53
                  Left            =   8260
                  TabIndex        =   874
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   53
                  Left            =   8660
                  TabIndex        =   873
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   53
                  Left            =   9050
                  TabIndex        =   872
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   53
                  Left            =   9250
                  TabIndex        =   871
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   53
                  Left            =   9460
                  TabIndex        =   870
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   53
                  Left            =   9660
                  TabIndex        =   869
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   53
                  Left            =   11085
                  TabIndex        =   868
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   53
                  Left            =   4850
                  TabIndex        =   867
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   53
                  Left            =   5240
                  TabIndex        =   866
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   53
                  Left            =   10870
                  TabIndex        =   865
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   53
                  Left            =   0
                  TabIndex        =   879
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   52
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   848
               Top             =   15600
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   52
                  Left            =   5240
                  TabIndex        =   862
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   52
                  Left            =   4850
                  TabIndex        =   861
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   52
                  Left            =   11085
                  TabIndex        =   860
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   52
                  Left            =   9660
                  TabIndex        =   859
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   52
                  Left            =   9460
                  TabIndex        =   858
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   52
                  Left            =   9250
                  TabIndex        =   857
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   52
                  Left            =   9050
                  TabIndex        =   856
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   52
                  Left            =   8660
                  TabIndex        =   855
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   52
                  Left            =   8260
                  TabIndex        =   854
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   52
                  Left            =   4460
                  TabIndex        =   853
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   52
                  Left            =   4250
                  TabIndex        =   852
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   52
                  Left            =   4050
                  TabIndex        =   851
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   52
                  Left            =   320
                  TabIndex        =   850
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   52
                  Left            =   10870
                  TabIndex        =   849
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   52
                  Left            =   0
                  TabIndex        =   863
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   51
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   832
               Top             =   15300
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   51
                  Left            =   10870
                  TabIndex        =   846
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   51
                  Left            =   5240
                  TabIndex        =   845
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   51
                  Left            =   4850
                  TabIndex        =   844
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   51
                  Left            =   11085
                  TabIndex        =   843
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   51
                  Left            =   9660
                  TabIndex        =   842
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   51
                  Left            =   9460
                  TabIndex        =   841
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   51
                  Left            =   9250
                  TabIndex        =   840
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   51
                  Left            =   9050
                  TabIndex        =   839
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   51
                  Left            =   8660
                  TabIndex        =   838
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   51
                  Left            =   8260
                  TabIndex        =   837
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   51
                  Left            =   4460
                  TabIndex        =   836
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   51
                  Left            =   4250
                  TabIndex        =   835
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   51
                  Left            =   4050
                  TabIndex        =   834
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   51
                  Left            =   320
                  TabIndex        =   833
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   51
                  Left            =   0
                  TabIndex        =   847
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   50
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   816
               Top             =   15000
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   50
                  Left            =   10870
                  TabIndex        =   830
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   50
                  Left            =   320
                  TabIndex        =   829
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   50
                  Left            =   4050
                  TabIndex        =   828
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   50
                  Left            =   4250
                  TabIndex        =   827
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   50
                  Left            =   4460
                  TabIndex        =   826
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   50
                  Left            =   8260
                  TabIndex        =   825
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   50
                  Left            =   8660
                  TabIndex        =   824
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   50
                  Left            =   9050
                  TabIndex        =   823
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   50
                  Left            =   9250
                  TabIndex        =   822
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   50
                  Left            =   9460
                  TabIndex        =   821
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   50
                  Left            =   9660
                  TabIndex        =   820
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   50
                  Left            =   11085
                  TabIndex        =   819
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   50
                  Left            =   4850
                  TabIndex        =   818
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   50
                  Left            =   5240
                  TabIndex        =   817
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   50
                  Left            =   0
                  TabIndex        =   831
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   49
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   800
               Top             =   14700
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   49
                  Left            =   320
                  TabIndex        =   814
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   49
                  Left            =   4050
                  TabIndex        =   813
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   49
                  Left            =   4250
                  TabIndex        =   812
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   49
                  Left            =   4460
                  TabIndex        =   811
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   49
                  Left            =   8260
                  TabIndex        =   810
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   49
                  Left            =   8660
                  TabIndex        =   809
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   49
                  Left            =   9050
                  TabIndex        =   808
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   49
                  Left            =   9250
                  TabIndex        =   807
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   49
                  Left            =   9460
                  TabIndex        =   806
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   49
                  Left            =   9660
                  TabIndex        =   805
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   49
                  Left            =   11085
                  TabIndex        =   804
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   49
                  Left            =   4850
                  TabIndex        =   803
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   49
                  Left            =   5240
                  TabIndex        =   802
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   49
                  Left            =   10870
                  TabIndex        =   801
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   49
                  Left            =   0
                  TabIndex        =   815
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   48
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   784
               Top             =   14400
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   48
                  Left            =   5240
                  TabIndex        =   798
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   48
                  Left            =   4850
                  TabIndex        =   797
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   48
                  Left            =   11085
                  TabIndex        =   796
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   48
                  Left            =   9660
                  TabIndex        =   795
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   48
                  Left            =   9460
                  TabIndex        =   794
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   48
                  Left            =   9250
                  TabIndex        =   793
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   48
                  Left            =   9050
                  TabIndex        =   792
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   48
                  Left            =   8660
                  TabIndex        =   791
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   48
                  Left            =   8260
                  TabIndex        =   790
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   48
                  Left            =   4460
                  TabIndex        =   789
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   48
                  Left            =   4250
                  TabIndex        =   788
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   48
                  Left            =   4050
                  TabIndex        =   787
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   48
                  Left            =   320
                  TabIndex        =   786
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   48
                  Left            =   10870
                  TabIndex        =   785
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   48
                  Left            =   0
                  TabIndex        =   799
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   47
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   768
               Top             =   14100
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   47
                  Left            =   320
                  TabIndex        =   782
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   47
                  Left            =   4050
                  TabIndex        =   781
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   47
                  Left            =   4250
                  TabIndex        =   780
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   47
                  Left            =   4460
                  TabIndex        =   779
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   47
                  Left            =   8260
                  TabIndex        =   778
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   47
                  Left            =   8660
                  TabIndex        =   777
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   47
                  Left            =   9050
                  TabIndex        =   776
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   47
                  Left            =   9250
                  TabIndex        =   775
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   47
                  Left            =   9460
                  TabIndex        =   774
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   47
                  Left            =   9660
                  TabIndex        =   773
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   47
                  Left            =   11085
                  TabIndex        =   772
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   47
                  Left            =   4850
                  TabIndex        =   771
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   47
                  Left            =   5240
                  TabIndex        =   770
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   47
                  Left            =   10870
                  TabIndex        =   769
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   47
                  Left            =   0
                  TabIndex        =   783
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   46
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   752
               Top             =   13800
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   46
                  Left            =   5240
                  TabIndex        =   766
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   46
                  Left            =   4850
                  TabIndex        =   765
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   46
                  Left            =   11085
                  TabIndex        =   764
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   46
                  Left            =   9660
                  TabIndex        =   763
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   46
                  Left            =   9460
                  TabIndex        =   762
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   46
                  Left            =   9250
                  TabIndex        =   761
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   46
                  Left            =   9050
                  TabIndex        =   760
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   46
                  Left            =   8660
                  TabIndex        =   759
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   46
                  Left            =   8260
                  TabIndex        =   758
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   46
                  Left            =   4460
                  TabIndex        =   757
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   46
                  Left            =   4250
                  TabIndex        =   756
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   46
                  Left            =   4050
                  TabIndex        =   755
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   46
                  Left            =   320
                  TabIndex        =   754
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   46
                  Left            =   10870
                  TabIndex        =   753
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   46
                  Left            =   0
                  TabIndex        =   767
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   45
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   736
               Top             =   13500
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   45
                  Left            =   10870
                  TabIndex        =   750
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   45
                  Left            =   5240
                  TabIndex        =   749
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   45
                  Left            =   4850
                  TabIndex        =   748
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   45
                  Left            =   11085
                  TabIndex        =   747
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   45
                  Left            =   9660
                  TabIndex        =   746
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   45
                  Left            =   9460
                  TabIndex        =   745
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   45
                  Left            =   9250
                  TabIndex        =   744
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   45
                  Left            =   9050
                  TabIndex        =   743
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   45
                  Left            =   8660
                  TabIndex        =   742
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   45
                  Left            =   8260
                  TabIndex        =   741
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   45
                  Left            =   4460
                  TabIndex        =   740
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   45
                  Left            =   4250
                  TabIndex        =   739
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   45
                  Left            =   4050
                  TabIndex        =   738
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   45
                  Left            =   320
                  TabIndex        =   737
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   45
                  Left            =   0
                  TabIndex        =   751
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   44
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   720
               Top             =   13200
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   44
                  Left            =   10870
                  TabIndex        =   734
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   44
                  Left            =   320
                  TabIndex        =   733
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   44
                  Left            =   4050
                  TabIndex        =   732
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   44
                  Left            =   4250
                  TabIndex        =   731
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   44
                  Left            =   4460
                  TabIndex        =   730
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   44
                  Left            =   8260
                  TabIndex        =   729
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   44
                  Left            =   8660
                  TabIndex        =   728
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   44
                  Left            =   9050
                  TabIndex        =   727
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   44
                  Left            =   9250
                  TabIndex        =   726
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   44
                  Left            =   9460
                  TabIndex        =   725
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   44
                  Left            =   9660
                  TabIndex        =   724
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   44
                  Left            =   11085
                  TabIndex        =   723
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   44
                  Left            =   4850
                  TabIndex        =   722
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   44
                  Left            =   5240
                  TabIndex        =   721
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   44
                  Left            =   0
                  TabIndex        =   735
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   43
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   704
               Top             =   12900
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   43
                  Left            =   320
                  TabIndex        =   718
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   43
                  Left            =   4050
                  TabIndex        =   717
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   43
                  Left            =   4250
                  TabIndex        =   716
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   43
                  Left            =   4460
                  TabIndex        =   715
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   43
                  Left            =   8260
                  TabIndex        =   714
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   43
                  Left            =   8660
                  TabIndex        =   713
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   43
                  Left            =   9050
                  TabIndex        =   712
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   43
                  Left            =   9250
                  TabIndex        =   711
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   43
                  Left            =   9460
                  TabIndex        =   710
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   43
                  Left            =   9660
                  TabIndex        =   709
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   43
                  Left            =   11085
                  TabIndex        =   708
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   43
                  Left            =   4850
                  TabIndex        =   707
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   43
                  Left            =   5240
                  TabIndex        =   706
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   43
                  Left            =   10870
                  TabIndex        =   705
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   43
                  Left            =   0
                  TabIndex        =   719
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   42
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   688
               Top             =   12600
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   42
                  Left            =   5240
                  TabIndex        =   702
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   42
                  Left            =   4850
                  TabIndex        =   701
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   42
                  Left            =   11085
                  TabIndex        =   700
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   42
                  Left            =   9660
                  TabIndex        =   699
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   42
                  Left            =   9460
                  TabIndex        =   698
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   42
                  Left            =   9250
                  TabIndex        =   697
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   42
                  Left            =   9050
                  TabIndex        =   696
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   42
                  Left            =   8660
                  TabIndex        =   695
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   42
                  Left            =   8260
                  TabIndex        =   694
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   42
                  Left            =   4460
                  TabIndex        =   693
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   42
                  Left            =   4250
                  TabIndex        =   692
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   42
                  Left            =   4050
                  TabIndex        =   691
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   42
                  Left            =   320
                  TabIndex        =   690
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   42
                  Left            =   10870
                  TabIndex        =   689
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   42
                  Left            =   0
                  TabIndex        =   703
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   41
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   672
               Top             =   12300
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   41
                  Left            =   320
                  TabIndex        =   686
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   41
                  Left            =   4050
                  TabIndex        =   685
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   41
                  Left            =   4250
                  TabIndex        =   684
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   41
                  Left            =   4460
                  TabIndex        =   683
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   41
                  Left            =   8260
                  TabIndex        =   682
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   41
                  Left            =   8660
                  TabIndex        =   681
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   41
                  Left            =   9050
                  TabIndex        =   680
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   41
                  Left            =   9250
                  TabIndex        =   679
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   41
                  Left            =   9460
                  TabIndex        =   678
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   41
                  Left            =   9660
                  TabIndex        =   677
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   41
                  Left            =   11085
                  TabIndex        =   676
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   41
                  Left            =   4850
                  TabIndex        =   675
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   41
                  Left            =   5240
                  TabIndex        =   674
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   41
                  Left            =   10870
                  TabIndex        =   673
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   41
                  Left            =   0
                  TabIndex        =   687
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   40
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   656
               Top             =   12000
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   40
                  Left            =   5240
                  TabIndex        =   670
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   40
                  Left            =   4850
                  TabIndex        =   669
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   40
                  Left            =   11085
                  TabIndex        =   668
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   40
                  Left            =   9660
                  TabIndex        =   667
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   40
                  Left            =   9460
                  TabIndex        =   666
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   40
                  Left            =   9250
                  TabIndex        =   665
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   40
                  Left            =   9050
                  TabIndex        =   664
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   40
                  Left            =   8660
                  TabIndex        =   663
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   40
                  Left            =   8260
                  TabIndex        =   662
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   40
                  Left            =   4460
                  TabIndex        =   661
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   40
                  Left            =   4250
                  TabIndex        =   660
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   40
                  Left            =   4050
                  TabIndex        =   659
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   40
                  Left            =   320
                  TabIndex        =   658
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   40
                  Left            =   10870
                  TabIndex        =   657
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   40
                  Left            =   0
                  TabIndex        =   671
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   39
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   640
               Top             =   11700
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   39
                  Left            =   320
                  TabIndex        =   654
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   39
                  Left            =   4050
                  TabIndex        =   653
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   39
                  Left            =   4250
                  TabIndex        =   652
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   39
                  Left            =   4460
                  TabIndex        =   651
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   39
                  Left            =   8260
                  TabIndex        =   650
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   39
                  Left            =   8660
                  TabIndex        =   649
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   39
                  Left            =   9050
                  TabIndex        =   648
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   39
                  Left            =   9250
                  TabIndex        =   647
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   39
                  Left            =   9460
                  TabIndex        =   646
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   39
                  Left            =   9660
                  TabIndex        =   645
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   39
                  Left            =   11085
                  TabIndex        =   644
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   39
                  Left            =   4850
                  TabIndex        =   643
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   39
                  Left            =   5240
                  TabIndex        =   642
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   39
                  Left            =   10870
                  TabIndex        =   641
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   39
                  Left            =   0
                  TabIndex        =   655
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   38
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   624
               Top             =   11400
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   38
                  Left            =   5240
                  TabIndex        =   638
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   38
                  Left            =   4850
                  TabIndex        =   637
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   38
                  Left            =   11085
                  TabIndex        =   636
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   38
                  Left            =   9660
                  TabIndex        =   635
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   38
                  Left            =   9460
                  TabIndex        =   634
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   38
                  Left            =   9250
                  TabIndex        =   633
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   38
                  Left            =   9050
                  TabIndex        =   632
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   38
                  Left            =   8660
                  TabIndex        =   631
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   38
                  Left            =   8260
                  TabIndex        =   630
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   38
                  Left            =   4460
                  TabIndex        =   629
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   38
                  Left            =   4250
                  TabIndex        =   628
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   38
                  Left            =   4050
                  TabIndex        =   627
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   38
                  Left            =   320
                  TabIndex        =   626
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   38
                  Left            =   10870
                  TabIndex        =   625
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   38
                  Left            =   0
                  TabIndex        =   639
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   37
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   608
               Top             =   11100
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   37
                  Left            =   10870
                  TabIndex        =   622
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   37
                  Left            =   5240
                  TabIndex        =   621
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   37
                  Left            =   4850
                  TabIndex        =   620
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   37
                  Left            =   11085
                  TabIndex        =   619
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   37
                  Left            =   9660
                  TabIndex        =   618
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   37
                  Left            =   9460
                  TabIndex        =   617
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   37
                  Left            =   9250
                  TabIndex        =   616
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   37
                  Left            =   9050
                  TabIndex        =   615
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   37
                  Left            =   8660
                  TabIndex        =   614
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   37
                  Left            =   8260
                  TabIndex        =   613
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   37
                  Left            =   4460
                  TabIndex        =   612
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   37
                  Left            =   4250
                  TabIndex        =   611
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   37
                  Left            =   4050
                  TabIndex        =   610
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   37
                  Left            =   320
                  TabIndex        =   609
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   37
                  Left            =   0
                  TabIndex        =   623
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   36
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   592
               Top             =   10800
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   36
                  Left            =   10870
                  TabIndex        =   606
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   36
                  Left            =   320
                  TabIndex        =   605
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   36
                  Left            =   4050
                  TabIndex        =   604
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   36
                  Left            =   4250
                  TabIndex        =   603
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   36
                  Left            =   4460
                  TabIndex        =   602
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   36
                  Left            =   8260
                  TabIndex        =   601
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   36
                  Left            =   8660
                  TabIndex        =   600
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   36
                  Left            =   9050
                  TabIndex        =   599
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   36
                  Left            =   9250
                  TabIndex        =   598
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   36
                  Left            =   9460
                  TabIndex        =   597
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   36
                  Left            =   9660
                  TabIndex        =   596
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   36
                  Left            =   11085
                  TabIndex        =   595
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   36
                  Left            =   4850
                  TabIndex        =   594
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   36
                  Left            =   5240
                  TabIndex        =   593
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   36
                  Left            =   0
                  TabIndex        =   607
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   35
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   576
               Top             =   10500
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   35
                  Left            =   10870
                  TabIndex        =   590
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   35
                  Left            =   5240
                  TabIndex        =   589
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   35
                  Left            =   4850
                  TabIndex        =   588
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   35
                  Left            =   11085
                  TabIndex        =   587
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   35
                  Left            =   9660
                  TabIndex        =   586
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   35
                  Left            =   9460
                  TabIndex        =   585
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   35
                  Left            =   9250
                  TabIndex        =   584
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   35
                  Left            =   9050
                  TabIndex        =   583
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   35
                  Left            =   8660
                  TabIndex        =   582
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   35
                  Left            =   8260
                  TabIndex        =   581
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   35
                  Left            =   4460
                  TabIndex        =   580
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   35
                  Left            =   4250
                  TabIndex        =   579
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   35
                  Left            =   4050
                  TabIndex        =   578
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   35
                  Left            =   320
                  TabIndex        =   577
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   35
                  Left            =   0
                  TabIndex        =   591
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   34
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   560
               Top             =   10200
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   34
                  Left            =   10870
                  TabIndex        =   574
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   34
                  Left            =   320
                  TabIndex        =   573
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   34
                  Left            =   4050
                  TabIndex        =   572
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   34
                  Left            =   4250
                  TabIndex        =   571
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   34
                  Left            =   4460
                  TabIndex        =   570
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   34
                  Left            =   8260
                  TabIndex        =   569
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   34
                  Left            =   8660
                  TabIndex        =   568
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   34
                  Left            =   9050
                  TabIndex        =   567
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   34
                  Left            =   9250
                  TabIndex        =   566
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   34
                  Left            =   9460
                  TabIndex        =   565
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   34
                  Left            =   9660
                  TabIndex        =   564
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   34
                  Left            =   11085
                  TabIndex        =   563
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   34
                  Left            =   4850
                  TabIndex        =   562
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   34
                  Left            =   5240
                  TabIndex        =   561
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   34
                  Left            =   0
                  TabIndex        =   575
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   33
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   544
               Top             =   9900
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   33
                  Left            =   10870
                  TabIndex        =   558
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   33
                  Left            =   5240
                  TabIndex        =   557
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   33
                  Left            =   4850
                  TabIndex        =   556
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   33
                  Left            =   11085
                  TabIndex        =   555
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   33
                  Left            =   9660
                  TabIndex        =   554
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   33
                  Left            =   9460
                  TabIndex        =   553
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   33
                  Left            =   9250
                  TabIndex        =   552
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   33
                  Left            =   9050
                  TabIndex        =   551
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   33
                  Left            =   8660
                  TabIndex        =   550
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   33
                  Left            =   8260
                  TabIndex        =   549
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   33
                  Left            =   4460
                  TabIndex        =   548
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   33
                  Left            =   4250
                  TabIndex        =   547
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   33
                  Left            =   4050
                  TabIndex        =   546
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   33
                  Left            =   320
                  TabIndex        =   545
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   33
                  Left            =   0
                  TabIndex        =   559
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   32
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   528
               Top             =   9600
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   32
                  Left            =   10870
                  TabIndex        =   542
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   32
                  Left            =   320
                  TabIndex        =   541
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   32
                  Left            =   4050
                  TabIndex        =   540
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   32
                  Left            =   4250
                  TabIndex        =   539
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   32
                  Left            =   4460
                  TabIndex        =   538
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   32
                  Left            =   8260
                  TabIndex        =   537
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   32
                  Left            =   8660
                  TabIndex        =   536
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   32
                  Left            =   9050
                  TabIndex        =   535
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   32
                  Left            =   9250
                  TabIndex        =   534
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   32
                  Left            =   9460
                  TabIndex        =   533
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   32
                  Left            =   9660
                  TabIndex        =   532
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   32
                  Left            =   11085
                  TabIndex        =   531
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   32
                  Left            =   4850
                  TabIndex        =   530
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   32
                  Left            =   5240
                  TabIndex        =   529
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   32
                  Left            =   0
                  TabIndex        =   543
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   31
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   512
               Top             =   9300
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   31
                  Left            =   320
                  TabIndex        =   526
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   31
                  Left            =   4050
                  TabIndex        =   525
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   31
                  Left            =   4250
                  TabIndex        =   524
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   31
                  Left            =   4460
                  TabIndex        =   523
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   31
                  Left            =   8260
                  TabIndex        =   522
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   31
                  Left            =   8660
                  TabIndex        =   521
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   31
                  Left            =   9050
                  TabIndex        =   520
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   31
                  Left            =   9250
                  TabIndex        =   519
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   31
                  Left            =   9460
                  TabIndex        =   518
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   31
                  Left            =   9660
                  TabIndex        =   517
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   31
                  Left            =   11085
                  TabIndex        =   516
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   31
                  Left            =   4850
                  TabIndex        =   515
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   31
                  Left            =   5240
                  TabIndex        =   514
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   31
                  Left            =   10870
                  TabIndex        =   513
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   31
                  Left            =   0
                  TabIndex        =   527
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   30
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   496
               Top             =   9000
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   30
                  Left            =   5240
                  TabIndex        =   510
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   30
                  Left            =   4850
                  TabIndex        =   509
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   30
                  Left            =   11085
                  TabIndex        =   508
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   30
                  Left            =   9660
                  TabIndex        =   507
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   30
                  Left            =   9460
                  TabIndex        =   506
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   30
                  Left            =   9250
                  TabIndex        =   505
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   30
                  Left            =   9050
                  TabIndex        =   504
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   30
                  Left            =   8660
                  TabIndex        =   503
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   30
                  Left            =   8260
                  TabIndex        =   502
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   30
                  Left            =   4460
                  TabIndex        =   501
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   30
                  Left            =   4250
                  TabIndex        =   500
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   30
                  Left            =   4050
                  TabIndex        =   499
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   30
                  Left            =   320
                  TabIndex        =   498
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   30
                  Left            =   10870
                  TabIndex        =   497
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   30
                  Left            =   0
                  TabIndex        =   511
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   29
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   480
               Top             =   8700
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   29
                  Left            =   320
                  TabIndex        =   494
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   29
                  Left            =   4050
                  TabIndex        =   493
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   29
                  Left            =   4250
                  TabIndex        =   492
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   29
                  Left            =   4460
                  TabIndex        =   491
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   29
                  Left            =   8260
                  TabIndex        =   490
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   29
                  Left            =   8660
                  TabIndex        =   489
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   29
                  Left            =   9050
                  TabIndex        =   488
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   29
                  Left            =   9250
                  TabIndex        =   487
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   29
                  Left            =   9460
                  TabIndex        =   486
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   29
                  Left            =   9660
                  TabIndex        =   485
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   29
                  Left            =   11085
                  TabIndex        =   484
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   29
                  Left            =   4850
                  TabIndex        =   483
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   29
                  Left            =   5240
                  TabIndex        =   482
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   29
                  Left            =   10870
                  TabIndex        =   481
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   29
                  Left            =   0
                  TabIndex        =   495
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   28
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   464
               Top             =   8400
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   28
                  Left            =   5240
                  TabIndex        =   478
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   28
                  Left            =   4850
                  TabIndex        =   477
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   28
                  Left            =   11085
                  TabIndex        =   476
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   28
                  Left            =   9660
                  TabIndex        =   475
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   28
                  Left            =   9460
                  TabIndex        =   474
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   28
                  Left            =   9250
                  TabIndex        =   473
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   28
                  Left            =   9050
                  TabIndex        =   472
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   28
                  Left            =   8660
                  TabIndex        =   471
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   28
                  Left            =   8260
                  TabIndex        =   470
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   28
                  Left            =   4460
                  TabIndex        =   469
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   28
                  Left            =   4250
                  TabIndex        =   468
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   28
                  Left            =   4050
                  TabIndex        =   467
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   28
                  Left            =   320
                  TabIndex        =   466
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   28
                  Left            =   10870
                  TabIndex        =   465
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   28
                  Left            =   0
                  TabIndex        =   479
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   27
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   448
               Top             =   8100
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   27
                  Left            =   10870
                  TabIndex        =   462
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   27
                  Left            =   5240
                  TabIndex        =   461
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   27
                  Left            =   4850
                  TabIndex        =   460
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   27
                  Left            =   11085
                  TabIndex        =   459
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   27
                  Left            =   9660
                  TabIndex        =   458
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   27
                  Left            =   9460
                  TabIndex        =   457
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   27
                  Left            =   9250
                  TabIndex        =   456
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   27
                  Left            =   9050
                  TabIndex        =   455
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   27
                  Left            =   8660
                  TabIndex        =   454
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   27
                  Left            =   8260
                  TabIndex        =   453
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   27
                  Left            =   4460
                  TabIndex        =   452
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   27
                  Left            =   4250
                  TabIndex        =   451
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   27
                  Left            =   4050
                  TabIndex        =   450
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   27
                  Left            =   320
                  TabIndex        =   449
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   27
                  Left            =   0
                  TabIndex        =   463
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   26
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   432
               Top             =   7800
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   26
                  Left            =   10870
                  TabIndex        =   446
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   26
                  Left            =   320
                  TabIndex        =   445
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   26
                  Left            =   4050
                  TabIndex        =   444
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   26
                  Left            =   4250
                  TabIndex        =   443
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   26
                  Left            =   4460
                  TabIndex        =   442
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   26
                  Left            =   8260
                  TabIndex        =   441
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   26
                  Left            =   8660
                  TabIndex        =   440
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   26
                  Left            =   9050
                  TabIndex        =   439
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   26
                  Left            =   9250
                  TabIndex        =   438
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   26
                  Left            =   9460
                  TabIndex        =   437
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   26
                  Left            =   9660
                  TabIndex        =   436
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   26
                  Left            =   11085
                  TabIndex        =   435
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   26
                  Left            =   4850
                  TabIndex        =   434
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   26
                  Left            =   5240
                  TabIndex        =   433
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   26
                  Left            =   0
                  TabIndex        =   447
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   25
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   416
               Top             =   7500
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   25
                  Left            =   320
                  TabIndex        =   430
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   25
                  Left            =   4050
                  TabIndex        =   429
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   25
                  Left            =   4250
                  TabIndex        =   428
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   25
                  Left            =   4460
                  TabIndex        =   427
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   25
                  Left            =   8260
                  TabIndex        =   426
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   25
                  Left            =   8660
                  TabIndex        =   425
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   25
                  Left            =   9050
                  TabIndex        =   424
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   25
                  Left            =   9250
                  TabIndex        =   423
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   25
                  Left            =   9460
                  TabIndex        =   422
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   25
                  Left            =   9660
                  TabIndex        =   421
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   25
                  Left            =   11085
                  TabIndex        =   420
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   25
                  Left            =   4850
                  TabIndex        =   419
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   25
                  Left            =   5240
                  TabIndex        =   418
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   25
                  Left            =   10870
                  TabIndex        =   417
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   25
                  Left            =   0
                  TabIndex        =   431
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   24
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   400
               Top             =   7200
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   24
                  Left            =   5240
                  TabIndex        =   414
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   24
                  Left            =   4850
                  TabIndex        =   413
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   24
                  Left            =   11085
                  TabIndex        =   412
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   24
                  Left            =   9660
                  TabIndex        =   411
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   24
                  Left            =   9460
                  TabIndex        =   410
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   24
                  Left            =   9250
                  TabIndex        =   409
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   24
                  Left            =   9050
                  TabIndex        =   408
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   24
                  Left            =   8660
                  TabIndex        =   407
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   24
                  Left            =   8260
                  TabIndex        =   406
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   24
                  Left            =   4460
                  TabIndex        =   405
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   24
                  Left            =   4250
                  TabIndex        =   404
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   24
                  Left            =   4050
                  TabIndex        =   403
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   24
                  Left            =   320
                  TabIndex        =   402
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   24
                  Left            =   10870
                  TabIndex        =   401
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   24
                  Left            =   0
                  TabIndex        =   415
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   23
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   384
               Top             =   6900
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   23
                  Left            =   10870
                  TabIndex        =   398
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   23
                  Left            =   5240
                  TabIndex        =   397
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   23
                  Left            =   4850
                  TabIndex        =   396
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   23
                  Left            =   11085
                  TabIndex        =   395
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   23
                  Left            =   9660
                  TabIndex        =   394
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   23
                  Left            =   9460
                  TabIndex        =   393
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   23
                  Left            =   9250
                  TabIndex        =   392
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   23
                  Left            =   9050
                  TabIndex        =   391
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   23
                  Left            =   8660
                  TabIndex        =   390
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   23
                  Left            =   8260
                  TabIndex        =   389
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   23
                  Left            =   4460
                  TabIndex        =   388
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   23
                  Left            =   4250
                  TabIndex        =   387
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   23
                  Left            =   4050
                  TabIndex        =   386
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   23
                  Left            =   320
                  TabIndex        =   385
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   23
                  Left            =   0
                  TabIndex        =   399
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   22
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   368
               Top             =   6600
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   22
                  Left            =   10870
                  TabIndex        =   382
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   22
                  Left            =   320
                  TabIndex        =   381
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   22
                  Left            =   4050
                  TabIndex        =   380
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   22
                  Left            =   4250
                  TabIndex        =   379
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   22
                  Left            =   4460
                  TabIndex        =   378
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   22
                  Left            =   8260
                  TabIndex        =   377
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   22
                  Left            =   8660
                  TabIndex        =   376
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   22
                  Left            =   9050
                  TabIndex        =   375
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   22
                  Left            =   9250
                  TabIndex        =   374
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   22
                  Left            =   9460
                  TabIndex        =   373
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   22
                  Left            =   9660
                  TabIndex        =   372
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   22
                  Left            =   11085
                  TabIndex        =   371
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   22
                  Left            =   4850
                  TabIndex        =   370
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   22
                  Left            =   5240
                  TabIndex        =   369
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   22
                  Left            =   0
                  TabIndex        =   383
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   21
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   352
               Top             =   6300
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   21
                  Left            =   10870
                  TabIndex        =   366
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   21
                  Left            =   5240
                  TabIndex        =   365
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   21
                  Left            =   4850
                  TabIndex        =   364
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   21
                  Left            =   11085
                  TabIndex        =   363
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   21
                  Left            =   9660
                  TabIndex        =   362
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   21
                  Left            =   9460
                  TabIndex        =   361
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   21
                  Left            =   9250
                  TabIndex        =   360
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   21
                  Left            =   9050
                  TabIndex        =   359
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   21
                  Left            =   8660
                  TabIndex        =   358
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   21
                  Left            =   8260
                  TabIndex        =   357
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   21
                  Left            =   4460
                  TabIndex        =   356
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   21
                  Left            =   4250
                  TabIndex        =   355
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   21
                  Left            =   4050
                  TabIndex        =   354
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   21
                  Left            =   320
                  TabIndex        =   353
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   21
                  Left            =   0
                  TabIndex        =   367
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   20
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   336
               Top             =   6000
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   20
                  Left            =   10870
                  TabIndex        =   350
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   20
                  Left            =   320
                  TabIndex        =   349
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   20
                  Left            =   4050
                  TabIndex        =   348
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   20
                  Left            =   4250
                  TabIndex        =   347
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   20
                  Left            =   4460
                  TabIndex        =   346
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   20
                  Left            =   8260
                  TabIndex        =   345
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   20
                  Left            =   8660
                  TabIndex        =   344
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   20
                  Left            =   9050
                  TabIndex        =   343
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   20
                  Left            =   9250
                  TabIndex        =   342
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   20
                  Left            =   9460
                  TabIndex        =   341
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   20
                  Left            =   9660
                  TabIndex        =   340
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   20
                  Left            =   11085
                  TabIndex        =   339
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   20
                  Left            =   4850
                  TabIndex        =   338
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   20
                  Left            =   5240
                  TabIndex        =   337
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   20
                  Left            =   0
                  TabIndex        =   351
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   19
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   320
               Top             =   5700
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   19
                  Left            =   320
                  TabIndex        =   334
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   19
                  Left            =   4050
                  TabIndex        =   333
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   19
                  Left            =   4250
                  TabIndex        =   332
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   19
                  Left            =   4460
                  TabIndex        =   331
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   19
                  Left            =   8260
                  TabIndex        =   330
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   19
                  Left            =   8660
                  TabIndex        =   329
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   19
                  Left            =   9050
                  TabIndex        =   328
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   19
                  Left            =   9250
                  TabIndex        =   327
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   19
                  Left            =   9460
                  TabIndex        =   326
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   19
                  Left            =   9660
                  TabIndex        =   325
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   19
                  Left            =   11085
                  TabIndex        =   324
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   19
                  Left            =   4850
                  TabIndex        =   323
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   19
                  Left            =   5240
                  TabIndex        =   322
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   19
                  Left            =   10870
                  TabIndex        =   321
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   19
                  Left            =   0
                  TabIndex        =   335
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   18
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   304
               Top             =   5400
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   18
                  Left            =   5240
                  TabIndex        =   318
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   18
                  Left            =   4850
                  TabIndex        =   317
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   18
                  Left            =   11085
                  TabIndex        =   316
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   18
                  Left            =   9660
                  TabIndex        =   315
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   18
                  Left            =   9460
                  TabIndex        =   314
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   18
                  Left            =   9250
                  TabIndex        =   313
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   18
                  Left            =   9050
                  TabIndex        =   312
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   18
                  Left            =   8660
                  TabIndex        =   311
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   18
                  Left            =   8260
                  TabIndex        =   310
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   18
                  Left            =   4460
                  TabIndex        =   309
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   18
                  Left            =   4250
                  TabIndex        =   308
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   18
                  Left            =   4050
                  TabIndex        =   307
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   18
                  Left            =   320
                  TabIndex        =   306
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   18
                  Left            =   10870
                  TabIndex        =   305
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   18
                  Left            =   0
                  TabIndex        =   319
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   17
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   288
               Top             =   5100
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   17
                  Left            =   10870
                  TabIndex        =   302
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   17
                  Left            =   5240
                  TabIndex        =   301
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   17
                  Left            =   4850
                  TabIndex        =   300
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   17
                  Left            =   11085
                  TabIndex        =   299
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   17
                  Left            =   9660
                  TabIndex        =   298
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   17
                  Left            =   9460
                  TabIndex        =   297
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   17
                  Left            =   9250
                  TabIndex        =   296
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   17
                  Left            =   9050
                  TabIndex        =   295
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   17
                  Left            =   8660
                  TabIndex        =   294
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   17
                  Left            =   8260
                  TabIndex        =   293
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   17
                  Left            =   4460
                  TabIndex        =   292
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   17
                  Left            =   4250
                  TabIndex        =   291
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   17
                  Left            =   4050
                  TabIndex        =   290
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   17
                  Left            =   320
                  TabIndex        =   289
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   17
                  Left            =   0
                  TabIndex        =   303
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   16
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   272
               Top             =   4800
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   16
                  Left            =   10870
                  TabIndex        =   286
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   16
                  Left            =   320
                  TabIndex        =   285
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   16
                  Left            =   4050
                  TabIndex        =   284
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   16
                  Left            =   4250
                  TabIndex        =   283
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   16
                  Left            =   4460
                  TabIndex        =   282
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   16
                  Left            =   8260
                  TabIndex        =   281
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   16
                  Left            =   8660
                  TabIndex        =   280
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   16
                  Left            =   9050
                  TabIndex        =   279
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   16
                  Left            =   9250
                  TabIndex        =   278
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   16
                  Left            =   9460
                  TabIndex        =   277
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   16
                  Left            =   9660
                  TabIndex        =   276
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   16
                  Left            =   11085
                  TabIndex        =   275
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   16
                  Left            =   4850
                  TabIndex        =   274
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   16
                  Left            =   5240
                  TabIndex        =   273
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   16
                  Left            =   0
                  TabIndex        =   287
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   15
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   256
               Top             =   4500
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   15
                  Left            =   10870
                  TabIndex        =   270
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   15
                  Left            =   5240
                  TabIndex        =   269
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   15
                  Left            =   4850
                  TabIndex        =   268
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   15
                  Left            =   11085
                  TabIndex        =   267
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   15
                  Left            =   9660
                  TabIndex        =   266
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   15
                  Left            =   9460
                  TabIndex        =   265
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   15
                  Left            =   9250
                  TabIndex        =   264
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   15
                  Left            =   9050
                  TabIndex        =   263
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   15
                  Left            =   8660
                  TabIndex        =   262
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   15
                  Left            =   8260
                  TabIndex        =   261
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   15
                  Left            =   4460
                  TabIndex        =   260
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   15
                  Left            =   4250
                  TabIndex        =   259
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   15
                  Left            =   4050
                  TabIndex        =   258
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   15
                  Left            =   320
                  TabIndex        =   257
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   15
                  Left            =   0
                  TabIndex        =   271
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   14
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   240
               Top             =   4200
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   14
                  Left            =   10870
                  TabIndex        =   254
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   14
                  Left            =   320
                  TabIndex        =   253
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   14
                  Left            =   4050
                  TabIndex        =   252
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   14
                  Left            =   4250
                  TabIndex        =   251
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   14
                  Left            =   4460
                  TabIndex        =   250
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   14
                  Left            =   8260
                  TabIndex        =   249
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   14
                  Left            =   8660
                  TabIndex        =   248
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   14
                  Left            =   9050
                  TabIndex        =   247
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   14
                  Left            =   9250
                  TabIndex        =   246
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   14
                  Left            =   9460
                  TabIndex        =   245
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   14
                  Left            =   9660
                  TabIndex        =   244
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   14
                  Left            =   11085
                  TabIndex        =   243
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   14
                  Left            =   4850
                  TabIndex        =   242
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   14
                  Left            =   5240
                  TabIndex        =   241
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   14
                  Left            =   0
                  TabIndex        =   255
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   13
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   224
               Top             =   3900
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   13
                  Left            =   320
                  TabIndex        =   238
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   13
                  Left            =   4050
                  TabIndex        =   237
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   13
                  Left            =   4250
                  TabIndex        =   236
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   13
                  Left            =   4460
                  TabIndex        =   235
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   13
                  Left            =   8260
                  TabIndex        =   234
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   13
                  Left            =   8660
                  TabIndex        =   233
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   13
                  Left            =   9050
                  TabIndex        =   232
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   13
                  Left            =   9250
                  TabIndex        =   231
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   13
                  Left            =   9460
                  TabIndex        =   230
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   13
                  Left            =   9660
                  TabIndex        =   229
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   13
                  Left            =   11085
                  TabIndex        =   228
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   13
                  Left            =   4850
                  TabIndex        =   227
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   13
                  Left            =   5240
                  TabIndex        =   226
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   13
                  Left            =   10870
                  TabIndex        =   225
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   13
                  Left            =   0
                  TabIndex        =   239
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   12
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   208
               Top             =   3600
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   12
                  Left            =   5240
                  TabIndex        =   222
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   12
                  Left            =   4850
                  TabIndex        =   221
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   12
                  Left            =   11085
                  TabIndex        =   220
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   12
                  Left            =   9660
                  TabIndex        =   219
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   12
                  Left            =   9460
                  TabIndex        =   218
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   12
                  Left            =   9250
                  TabIndex        =   217
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   12
                  Left            =   9050
                  TabIndex        =   216
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   12
                  Left            =   8660
                  TabIndex        =   215
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   12
                  Left            =   8260
                  TabIndex        =   214
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   12
                  Left            =   4460
                  TabIndex        =   213
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   12
                  Left            =   4250
                  TabIndex        =   212
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   12
                  Left            =   4050
                  TabIndex        =   211
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   12
                  Left            =   320
                  TabIndex        =   210
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   12
                  Left            =   10870
                  TabIndex        =   209
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   12
                  Left            =   0
                  TabIndex        =   223
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   11
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   192
               Top             =   3300
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   11
                  Left            =   10870
                  TabIndex        =   206
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   11
                  Left            =   5240
                  TabIndex        =   205
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   11
                  Left            =   4850
                  TabIndex        =   204
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   11
                  Left            =   11085
                  TabIndex        =   203
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   11
                  Left            =   9660
                  TabIndex        =   202
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   11
                  Left            =   9460
                  TabIndex        =   201
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   11
                  Left            =   9250
                  TabIndex        =   200
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   11
                  Left            =   9050
                  TabIndex        =   199
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   11
                  Left            =   8660
                  TabIndex        =   198
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   11
                  Left            =   8260
                  TabIndex        =   197
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   11
                  Left            =   4460
                  TabIndex        =   196
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   11
                  Left            =   4250
                  TabIndex        =   195
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   11
                  Left            =   4050
                  TabIndex        =   194
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   11
                  Left            =   320
                  TabIndex        =   193
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   11
                  Left            =   0
                  TabIndex        =   207
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   10
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   176
               Top             =   3000
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   10
                  Left            =   10870
                  TabIndex        =   190
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   10
                  Left            =   320
                  TabIndex        =   189
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   10
                  Left            =   4050
                  TabIndex        =   188
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   10
                  Left            =   4250
                  TabIndex        =   187
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   10
                  Left            =   4460
                  TabIndex        =   186
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   10
                  Left            =   8260
                  TabIndex        =   185
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   10
                  Left            =   8660
                  TabIndex        =   184
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   10
                  Left            =   9050
                  TabIndex        =   183
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   10
                  Left            =   9250
                  TabIndex        =   182
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   10
                  Left            =   9460
                  TabIndex        =   181
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   10
                  Left            =   9660
                  TabIndex        =   180
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   10
                  Left            =   11085
                  TabIndex        =   179
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   10
                  Left            =   4850
                  TabIndex        =   178
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   10
                  Left            =   5240
                  TabIndex        =   177
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   10
                  Left            =   0
                  TabIndex        =   191
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   9
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   160
               Top             =   2700
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   9
                  Left            =   10870
                  TabIndex        =   174
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   9
                  Left            =   5240
                  TabIndex        =   173
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   9
                  Left            =   4850
                  TabIndex        =   172
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   9
                  Left            =   11085
                  TabIndex        =   171
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   9
                  Left            =   9660
                  TabIndex        =   170
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   9
                  Left            =   9460
                  TabIndex        =   169
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   9
                  Left            =   9250
                  TabIndex        =   168
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   9
                  Left            =   9050
                  TabIndex        =   167
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   9
                  Left            =   8660
                  TabIndex        =   166
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   9
                  Left            =   8260
                  TabIndex        =   165
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   9
                  Left            =   4460
                  TabIndex        =   164
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   9
                  Left            =   4250
                  TabIndex        =   163
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   9
                  Left            =   4050
                  TabIndex        =   162
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   9
                  Left            =   320
                  TabIndex        =   161
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   9
                  Left            =   0
                  TabIndex        =   175
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   8
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   144
               Top             =   2400
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   8
                  Left            =   10870
                  TabIndex        =   158
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   8
                  Left            =   320
                  TabIndex        =   157
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   8
                  Left            =   4050
                  TabIndex        =   156
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   8
                  Left            =   4250
                  TabIndex        =   155
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   8
                  Left            =   4460
                  TabIndex        =   154
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   8
                  Left            =   8260
                  TabIndex        =   153
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   8
                  Left            =   8660
                  TabIndex        =   152
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   8
                  Left            =   9050
                  TabIndex        =   151
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   8
                  Left            =   9250
                  TabIndex        =   150
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   8
                  Left            =   9460
                  TabIndex        =   149
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   8
                  Left            =   9660
                  TabIndex        =   148
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   8
                  Left            =   11085
                  TabIndex        =   147
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   8
                  Left            =   4850
                  TabIndex        =   146
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   8
                  Left            =   5240
                  TabIndex        =   145
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   8
                  Left            =   0
                  TabIndex        =   159
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   7
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   128
               Top             =   2100
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   7
                  Left            =   10870
                  TabIndex        =   142
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   7
                  Left            =   5240
                  TabIndex        =   141
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   7
                  Left            =   4850
                  TabIndex        =   140
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   7
                  Left            =   11085
                  TabIndex        =   139
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   7
                  Left            =   9660
                  TabIndex        =   138
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   7
                  Left            =   9460
                  TabIndex        =   137
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   7
                  Left            =   9250
                  TabIndex        =   136
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   7
                  Left            =   9050
                  TabIndex        =   135
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   7
                  Left            =   8660
                  TabIndex        =   134
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   7
                  Left            =   8260
                  TabIndex        =   133
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   7
                  Left            =   4460
                  TabIndex        =   132
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   7
                  Left            =   4250
                  TabIndex        =   131
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   7
                  Left            =   4050
                  TabIndex        =   130
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   7
                  Left            =   320
                  TabIndex        =   129
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   7
                  Left            =   0
                  TabIndex        =   143
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   6
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   112
               Top             =   1800
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   6
                  Left            =   10870
                  TabIndex        =   126
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   6
                  Left            =   320
                  TabIndex        =   125
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   6
                  Left            =   4050
                  TabIndex        =   124
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   6
                  Left            =   4250
                  TabIndex        =   123
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   6
                  Left            =   4460
                  TabIndex        =   122
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   6
                  Left            =   8260
                  TabIndex        =   121
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   6
                  Left            =   8660
                  TabIndex        =   120
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   6
                  Left            =   9050
                  TabIndex        =   119
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   6
                  Left            =   9250
                  TabIndex        =   118
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   6
                  Left            =   9460
                  TabIndex        =   117
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   6
                  Left            =   9660
                  TabIndex        =   116
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   6
                  Left            =   11085
                  TabIndex        =   115
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   6
                  Left            =   4850
                  TabIndex        =   114
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   6
                  Left            =   5240
                  TabIndex        =   113
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   6
                  Left            =   0
                  TabIndex        =   127
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   5
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   96
               Top             =   1500
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   5
                  Left            =   320
                  TabIndex        =   110
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   5
                  Left            =   4050
                  TabIndex        =   109
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   5
                  Left            =   4250
                  TabIndex        =   108
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   5
                  Left            =   4460
                  TabIndex        =   107
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   5
                  Left            =   8260
                  TabIndex        =   106
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   5
                  Left            =   8660
                  TabIndex        =   105
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   5
                  Left            =   9050
                  TabIndex        =   104
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   5
                  Left            =   9250
                  TabIndex        =   103
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   5
                  Left            =   9460
                  TabIndex        =   102
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   5
                  Left            =   9660
                  TabIndex        =   101
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   5
                  Left            =   11085
                  TabIndex        =   100
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   5
                  Left            =   4850
                  TabIndex        =   99
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   5
                  Left            =   5240
                  TabIndex        =   98
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   5
                  Left            =   10870
                  TabIndex        =   97
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   5
                  Left            =   0
                  TabIndex        =   111
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   4
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   80
               Top             =   1200
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   4
                  Left            =   5240
                  TabIndex        =   94
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   4
                  Left            =   4850
                  TabIndex        =   93
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   4
                  Left            =   11085
                  TabIndex        =   92
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   4
                  Left            =   9660
                  TabIndex        =   91
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   4
                  Left            =   9460
                  TabIndex        =   90
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   4
                  Left            =   9250
                  TabIndex        =   89
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   4
                  Left            =   9050
                  TabIndex        =   88
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   4
                  Left            =   8660
                  TabIndex        =   87
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   4
                  Left            =   8260
                  TabIndex        =   86
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   4
                  Left            =   4460
                  TabIndex        =   85
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   4
                  Left            =   4250
                  TabIndex        =   84
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   4
                  Left            =   4050
                  TabIndex        =   83
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   4
                  Left            =   320
                  TabIndex        =   82
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   4
                  Left            =   10870
                  TabIndex        =   81
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   4
                  Left            =   0
                  TabIndex        =   95
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   3
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   64
               Top             =   900
               Width           =   13000
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   3
                  Left            =   320
                  TabIndex        =   78
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   3
                  Left            =   4050
                  TabIndex        =   77
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   3
                  Left            =   4250
                  TabIndex        =   76
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   3
                  Left            =   4460
                  TabIndex        =   75
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   3
                  Left            =   8260
                  TabIndex        =   74
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   3
                  Left            =   8660
                  TabIndex        =   73
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   3
                  Left            =   9050
                  TabIndex        =   72
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   3
                  Left            =   9250
                  TabIndex        =   71
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   3
                  Left            =   9460
                  TabIndex        =   70
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   3
                  Left            =   9660
                  TabIndex        =   69
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   3
                  Left            =   11085
                  TabIndex        =   68
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   3
                  Left            =   4850
                  TabIndex        =   67
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   3
                  Left            =   5240
                  TabIndex        =   66
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   3
                  Left            =   10870
                  TabIndex        =   65
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   3
                  Left            =   0
                  TabIndex        =   79
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   2
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   48
               Top             =   600
               Width           =   13000
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   2
                  Left            =   5240
                  TabIndex        =   62
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   2
                  Left            =   4850
                  TabIndex        =   61
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   2
                  Left            =   11085
                  TabIndex        =   60
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   2
                  Left            =   9660
                  TabIndex        =   59
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   2
                  Left            =   9460
                  TabIndex        =   58
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   2
                  Left            =   9250
                  TabIndex        =   57
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   2
                  Left            =   9050
                  TabIndex        =   56
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   2
                  Left            =   8660
                  TabIndex        =   55
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   2
                  Left            =   8260
                  TabIndex        =   54
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   2
                  Left            =   4460
                  TabIndex        =   53
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   2
                  Left            =   4250
                  TabIndex        =   52
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   2
                  Left            =   4050
                  TabIndex        =   51
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   2
                  Left            =   320
                  TabIndex        =   50
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   2
                  Left            =   10870
                  TabIndex        =   49
                  Top             =   0
                  Width           =   175
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   2
                  Left            =   0
                  TabIndex        =   63
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   0
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   32
               Top             =   0
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   0
                  Left            =   10870
                  TabIndex        =   46
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   0
                  Left            =   5240
                  TabIndex        =   45
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   0
                  Left            =   4850
                  TabIndex        =   44
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   0
                  Left            =   11085
                  TabIndex        =   43
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   0
                  Left            =   9660
                  TabIndex        =   42
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   0
                  Left            =   9460
                  TabIndex        =   41
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   0
                  Left            =   9250
                  TabIndex        =   40
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   0
                  Left            =   9050
                  TabIndex        =   39
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   0
                  Left            =   8660
                  TabIndex        =   38
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   0
                  Left            =   8260
                  TabIndex        =   37
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   0
                  Left            =   4460
                  TabIndex        =   36
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   0
                  Left            =   4250
                  TabIndex        =   35
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   0
                  Left            =   4050
                  TabIndex        =   34
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   0
                  Left            =   320
                  TabIndex        =   33
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   0
                  Left            =   0
                  TabIndex        =   47
                  Top             =   15
                  Width           =   300
               End
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   1
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   13005
               TabIndex        =   16
               Top             =   300
               Width           =   13000
               Begin VB.CommandButton Command10 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   1
                  Left            =   10870
                  TabIndex        =   30
                  Top             =   0
                  Width           =   175
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Index           =   1
                  Left            =   320
                  TabIndex        =   29
                  Top             =   0
                  Width           =   3720
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Index           =   1
                  Left            =   4050
                  TabIndex        =   28
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text3 
                  Height          =   285
                  Index           =   1
                  Left            =   4250
                  TabIndex        =   27
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   1
                  Left            =   4460
                  TabIndex        =   26
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   1
                  Left            =   8260
                  TabIndex        =   25
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox TexT8 
                  Height          =   285
                  Index           =   1
                  Left            =   8660
                  TabIndex        =   24
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text9 
                  Height          =   285
                  Index           =   1
                  Left            =   9050
                  TabIndex        =   23
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   1
                  Left            =   9250
                  TabIndex        =   22
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text11 
                  Height          =   285
                  Index           =   1
                  Left            =   9460
                  TabIndex        =   21
                  Top             =   0
                  Width           =   200
               End
               Begin VB.TextBox Text12 
                  Height          =   285
                  Index           =   1
                  Left            =   9660
                  TabIndex        =   20
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Index           =   1
                  Left            =   11085
                  TabIndex        =   19
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   1
                  Left            =   4850
                  TabIndex        =   18
                  Top             =   0
                  Width           =   370
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   1
                  Left            =   5240
                  TabIndex        =   17
                  Top             =   0
                  Width           =   3015
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   1
                  Left            =   0
                  TabIndex        =   31
                  Top             =   15
                  Width           =   300
               End
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   5400
            LargeChange     =   7
            Left            =   11520
            TabIndex        =   7
            Top             =   0
            Width           =   160
         End
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"PARAVEN2.frx":0000
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
         Left            =   450
         TabIndex        =   9
         Top             =   240
         Width           =   11415
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   960
      Left            =   10300
      ScaleHeight     =   900
      ScaleWidth      =   1965
      TabIndex        =   1
      Top             =   0
      Width           =   2025
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   860
         Picture         =   "PARAVEN2.frx":0097
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   150
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -320
         Picture         =   "PARAVEN2.frx":01E1
         Top             =   400
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   820
      Left            =   9480
      Picture         =   "PARAVEN2.frx":2103
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Guarda Configuración de Correo Electrónico"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label5 
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
      Left            =   4920
      TabIndex        =   15
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label4 
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
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label3 
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
      Left            =   7800
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Left            =   3240
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblServer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   11
      Top             =   180
      Width           =   720
   End
   Begin VB.Label lblFromName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6960
      TabIndex        =   10
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "PARAVEN2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim APA$

Private Sub Command1_Click()
   Unload Me
End Sub
'
Private Sub Command18_Click()
   Call ARMALISTARCH(APA$, "PRF", MKI$(6))    ' EL 6 ES POR VENTANA TABULADA
End Sub

Private Sub Command2_Click()
   PARAVENT2.Show 1
   
End Sub

Private Sub Form_Load()
   APA$ = PARAVENT.APA$
   '
   VScroll1.Max = (Frame15.Height - VScroll1.Height) / 100
   VScroll1.Value = 0
   '
   For KKI% = 1 To 64
      LEFE(KKI% - 1) = TRIM$(Str$(KKI%)) + "."
   Next KKI%
End Sub

Sub CREASTRUVEN()
   '
   Dim RGNX$(128)
   RGNX$(0) = String$(128, 0)
   TTXX$ = AJUSTI$("ESTRUCTURA TABLA RANDOM", 47) + Chr$(0) + Chr$(128)
   Call REPLA(RGNX$(0), TTXX$, 1, 49)
   '
   RGNX$(1) = ARMAREN("NOMBRE_DEL_CAMPO", 6, 0, 1, 32, 0)
   RGNX$(2) = ARMAREN("TIPO DE VARIABLE", 1, 0, 33, 1, 0)
   RGNX$(3) = ARMAREN("LONGITUD DECIMAL", 1, 0, 34, 1, 0)
   RGNX$(4) = ARMAREN("POSICION INICIAL", 1, 0, 35, 1, 0)
   RGNX$(5) = ARMAREN("LONGITUD CAMPO", 1, 0, 36, 1, 0)
   RGNX$(6) = ARMAREN("TEXTO ENCABEZADO", 6, 0, 37, 32, 0)
   RGNX$(7) = ARMAREN("ANCHO FISICO COLUMNA", 1, 0, 69, 1, 0)
   RGNX$(8) = ARMAREN("ANCHO LOGICO COLUMNA", 1, 0, 70, 1, 0)
   RGNX$(9) = ARMAREN("CARACTERES MARGEN", 1, 0, 71, 1, 0)
   RGNX$(10) = ARMAREN("JUSTIFICACION (1/2)", 1, 0, 72, 1, 0)
   RGNX$(11) = ARMAREN("EDITABLE (0/1/2)", 1, 0, 73, 1, 0)
   RGNX$(12) = ARMAREN("ARCHIVO ECO", 6, 0, 74, 8, 0)
   RGNX$(13) = ARMAREN("COLUMNA ECO", 1, 0, 82, 1, 0)
   '
   TTXX$ = ""
   For KKII% = 0 To 13
      TTXX$ = TTXX$ + RGNX$(KKII%)
   Next KKII%
   Call SAVEFILE(LUDAT$ + "STRUVEN.RFS", TTXX$)
   '
End Sub

Function ARMAREN$(NOCA$, TIVAX%, LODEC%, PINIX%, LOCAM%, Index%)
   '
   ARE$ = String$(128, 0)
   Call REPLA(ARE$, UCase$(AJUSTI$(NOCA$, 32)), 1, 32)
   Call REPLA(ARE$, Chr$(TIVAX%), 33, 1)
   Call REPLA(ARE$, Chr$(LODEC%), 34, 1)
   Call REPLA(ARE$, Chr$(PINIX%), 35, 1)
   Call REPLA(ARE$, Chr$(LOCAM%), 36, 1)
   Call REPLA(ARE$, Chr$(Index%), 37, 1)
   ARMAREN$ = ARE$
   '
End Function

Private Sub Label2_Change()
   If Label2 <> "" Then
      TTZZ$ = LOADFILE$(APA$ + "\" + Label2 + ".PRF")
      Label4 = UCase$(TRIM$(Left$(TTZZ$, 47)))
      JJ& = 0
      For KKII% = 1 To Len(TTZZ$) - 128 Step 128
        ZZZ$ = Mid$(TTZZ$, 256 + KKII%, 128)
        'If TRIM$(ZZZ$) <> "" Then
          Text6(JJ&) = TRIM$(Mid$(ZZZ$, 37, 32))
          Text7(JJ&) = TRIM$(Str$(Asc(Mid$(ZZZ$, 69, 1))))
          TexT8(JJ&) = TRIM$(Str$(Asc(Mid$(ZZZ$, 70, 1))))
          Text9(JJ&) = TRIM$(Str$(Asc(Mid$(ZZZ$, 71, 1))))
          Text10(JJ&) = TRIM$(Str$(Asc(Mid$(ZZZ$, 72, 1))))
          Text11(JJ&) = TRIM$(Str$(Asc(Mid$(ZZZ$, 73, 1))))
          Text12(JJ&) = TRIM$(Mid$(ZZZ$, 74, 8))
          Text13(JJ&) = TRIM$(Str$(Asc(Mid$(ZZZ$, 82, 1))))
          JJ& = JJ& + 1
        'End If
      Next KKII%
   End If
End Sub

Private Sub Label3_Change()
   If Label3 <> "" Then
      TTZZ$ = LOADFILE$(LUDAT$ + Label3 + ".RFS")
      Label5 = UCase$(TRIM$(Left$(TTZZ$, 48)))
      JJ& = 0
      For KKII% = 1 To Len(TTZZ$) - 128 Step 128
        ZZZ$ = Mid$(TTZZ$, 128 + KKII%, 128)
        If Len(ZZZ$) >= 128 Then
          Text1(JJ&) = TRIM$(Left$(ZZZ$, 32))
          Text2(JJ&) = TRIM$(Str$(Asc(Mid$(ZZZ$, 33, 1))))
          Text3(JJ&) = TRIM$(Str$(Asc(Mid$(ZZZ$, 34, 1))))
          Text4(JJ&) = TRIM$(Str$(Asc(Mid$(ZZZ$, 35, 1))))
          Text5(JJ&) = TRIM$(Str$(Asc(Mid$(ZZZ$, 36, 1))))
          JJ& = JJ& + 1
        End If
      Next KKII%
   End If
End Sub

Private Sub VScroll1_Change()
   Frame15.Top = (-1) * VScroll1.Value * 100
End Sub

Private Sub VScroll1_Scroll()
   Frame15.Top = (-1) * VScroll1.Value * 100
End Sub

