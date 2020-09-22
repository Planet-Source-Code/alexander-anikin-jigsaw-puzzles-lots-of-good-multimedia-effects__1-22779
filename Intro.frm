VERSION 5.00
Begin VB.Form Intro 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "Intro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Intro.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1800
      Top             =   1320
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5610
      Left            =   240
      Picture         =   "Intro.frx":0994
      ScaleHeight     =   5610
      ScaleWidth      =   7500
      TabIndex        =   0
      Top             =   840
      Width           =   7500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Load ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   360
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1065
   End
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'Copyright © 2001 by Alexander Anikin
'e-mail: aka@i.com.ua
'http://www.i.com.ua/~aka
'*************************************
Private Sub Form_Activate()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
 Unload Me
 Exit Sub
End If
Picture1.Left = (Screen.Width - Picture1.Width) \ 2
Picture1.Top = (Screen.Height - Picture1.Height) \ 2
Label2.Left = (Screen.Width - Label2.Width) \ 2
Label2.Top = Picture1.Top + Picture1.Height
End Sub


Private Sub Timer1_Timer()
Label2.Caption = "visit me: http://www.i.com.ua/~aka"
Label2.Left = (Screen.Width - Label2.Width) \ 2
Label2.Top = Picture1.Top + Picture1.Height
BeginPlaySound 63
Form1.Show
Timer1.Enabled = False
End Sub
