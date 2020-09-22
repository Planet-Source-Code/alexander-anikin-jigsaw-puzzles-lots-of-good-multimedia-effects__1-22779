VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Puzzles"
   ClientHeight    =   4860
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   324
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   2880
      Top             =   2160
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1800
      Top             =   2160
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   1800
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1800
      Top             =   2160
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Index           =   0
      Left            =   1800
      Picture         =   "Form1.frx":CD61
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   9600
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   1
         Left            =   0
         Picture         =   "Form1.frx":22674
         ScaleHeight     =   480
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   640
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   9600
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   7200
            Index           =   2
            Left            =   0
            Picture         =   "Form1.frx":3CB2C
            ScaleHeight     =   480
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   640
            TabIndex        =   5
            Top             =   600
            Visible         =   0   'False
            Width           =   9600
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   7200
               Index           =   3
               Left            =   0
               Picture         =   "Form1.frx":54E94
               ScaleHeight     =   480
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   640
               TabIndex        =   6
               Top             =   600
               Visible         =   0   'False
               Width           =   9600
               Begin VB.PictureBox Picture5 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H80000005&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   7200
                  Index           =   4
                  Left            =   0
                  Picture         =   "Form1.frx":67C89
                  ScaleHeight     =   480
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   640
                  TabIndex        =   7
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   9600
                  Begin VB.PictureBox Picture5 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     BackColor       =   &H80000005&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   7200
                     Index           =   5
                     Left            =   0
                     Picture         =   "Form1.frx":7C786
                     ScaleHeight     =   480
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   640
                     TabIndex        =   8
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   9600
                     Begin VB.PictureBox Picture5 
                        Appearance      =   0  'Flat
                        AutoRedraw      =   -1  'True
                        BackColor       =   &H80000005&
                        BorderStyle     =   0  'None
                        ForeColor       =   &H80000008&
                        Height          =   7200
                        Index           =   6
                        Left            =   120
                        Picture         =   "Form1.frx":93262
                        ScaleHeight     =   480
                        ScaleMode       =   3  'Pixel
                        ScaleWidth      =   640
                        TabIndex        =   9
                        Top             =   600
                        Visible         =   0   'False
                        Width           =   9600
                        Begin VB.PictureBox Picture5 
                           Appearance      =   0  'Flat
                           AutoRedraw      =   -1  'True
                           BackColor       =   &H80000005&
                           BorderStyle     =   0  'None
                           ForeColor       =   &H80000008&
                           Height          =   7200
                           Index           =   7
                           Left            =   0
                           Picture         =   "Form1.frx":A6FB9
                           ScaleHeight     =   480
                           ScaleMode       =   3  'Pixel
                           ScaleWidth      =   640
                           TabIndex        =   10
                           Top             =   480
                           Visible         =   0   'False
                           Width           =   9600
                           Begin VB.PictureBox Picture5 
                              Appearance      =   0  'Flat
                              AutoRedraw      =   -1  'True
                              BackColor       =   &H80000005&
                              BorderStyle     =   0  'None
                              ForeColor       =   &H80000008&
                              Height          =   7200
                              Index           =   8
                              Left            =   0
                              Picture         =   "Form1.frx":BDB9C
                              ScaleHeight     =   480
                              ScaleMode       =   3  'Pixel
                              ScaleWidth      =   640
                              TabIndex        =   11
                              Top             =   360
                              Visible         =   0   'False
                              Width           =   9600
                              Begin VB.PictureBox Picture5 
                                 Appearance      =   0  'Flat
                                 AutoRedraw      =   -1  'True
                                 BackColor       =   &H80000005&
                                 BorderStyle     =   0  'None
                                 ForeColor       =   &H80000008&
                                 Height          =   7200
                                 Index           =   9
                                 Left            =   0
                                 Picture         =   "Form1.frx":D5076
                                 ScaleHeight     =   480
                                 ScaleMode       =   3  'Pixel
                                 ScaleWidth      =   640
                                 TabIndex        =   12
                                 Top             =   360
                                 Visible         =   0   'False
                                 Width           =   9600
                              End
                           End
                        End
                     End
                  End
               End
            End
         End
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   0
      Picture         =   "Form1.frx":E9622
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   1
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7200
         Left            =   1200
         Picture         =   "Form1.frx":FE7EC
         ScaleHeight     =   480
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   640
         TabIndex        =   16
         Top             =   885
         Visible         =   0   'False
         Width           =   9600
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1200
         Index           =   48
         Left            =   1200
         ScaleHeight     =   80
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   80
         TabIndex        =   2
         Top             =   885
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LEVEL 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   2760
         TabIndex        =   14
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LEVEL 10 COMPLETE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   555
         Left            =   3000
         TabIndex        =   13
         Top             =   2400
         Width           =   5175
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   74
      Left            =   480
      Top             =   2160
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1260
      Left            =   1800
      Picture         =   "Form1.frx":113C72
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1260
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************
'Copyright © 2001 by Alexander Anikin
'e-mail: aka@i.com.ua
'http://www.i.com.ua/~aka
'***************************************************
'Shareware product!
'If you want to use this example of the code,
'you should send the check for 10 dollars by mail at
'    Alexander Anikin,
'    str. Verkhniaya 3, 74
'    Kiev, Ukraine, 01133.
'In the other case
'you should delete this files within 24 hours.
'***************************************************
'PS (Russian)
'Äëÿ âñåé òåððèòîðèè áûâøåãî ÑÑÑÐ áåñïëàòíî
'***************************************************

Dim i
Dim Dost As String
Dim UnDost
Dim z As Integer
Dim Num As Integer
Dim Pict As Picture
Dim pic1(47) As Integer
Private Sub Puzzle()
Dim imgX As ListImage
Static xxx As Integer
Static yyy As Integer
xxx = 0
yyy = 0

ImageList1.ListImages.Clear

For e = 0 To 47
  Picture2(e).PaintPicture Picture1.Picture, 0, 0, 80, 80, 80 * xxx, 80 * yyy, 80, 80
 xxx = xxx + 1: If xxx = 8 Then xxx = 0: yyy = yyy + 1
 Set imgX = ImageList1.ListImages. _
 Add(e + 1, , Picture2(e).Image)
 Intro.Label2.Caption = "Load data " & e
 DoEvents
Next e


 Erase pic1


For e = 0 To 47
 Picture2(e).Cls
upp:
 Randomize
 
 
 
 z = Rnd * 47
 If pic1(z) = 777 Then GoTo upp
 Picture2(e).Picture = ImageList1.ListImages(z + 1).ExtractIcon
 pic1(z) = 777
 Intro.Label2.Caption = "Load data " & e
 DoEvents
Next e

End Sub

Private Sub NextLevel()
Timer1.Enabled = False
Timer2.Enabled = False
If Num = 10 Then GoTo ex
Label1.Visible = False
Label1.Caption = "LEVEL " & (Num + 2)
Label1.Left = (Picture4.Width - Label1.Width) \ 2
Label2.Caption = "LEVEL " & (Num + 1) & " COMPLETE"
Label2.Left = (Picture4.Width - Label2.Width) \ 2
Label2.Top = (Picture4.Height - Label2.Height) \ 2

For e = 0 To 47
 Picture2(e).Visible = False
Next e
Picture1 = Picture5(Num)
Call Puzzle

For e = 0 To 47
 Picture2(e).Visible = True
Next e
Num = Num + 1
SaveSetting "Puzzles", "Startup", "Level", Num
Timer1.Enabled = True
Label1.Visible = True
UnDost = 0
Exit Sub
ex:
'''''
DeleteSetting "Puzzles", "Startup"
Label1.Visible = False
Label2.ForeColor = vbRed
Label2.Visible = False
MousePointer = 99
MouseIcon = Intro.MouseIcon

Picture4 = LoadPicture()
Picture4.PaintPicture Pict, 0, 0
Picture1 = LoadPicture()
Picture1.Visible = True
Picture1.PaintPicture Picture6, 0, 0
For e = 0 To 47
 Picture2(e).Visible = False
Next e
EndPlaySound
BeginPlaySound 65
Timer1.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Cheat()
If MousePointer = 0 Then
UnDost = 1
For e = 0 To 47
 Picture2(e) = ImageList1.ListImages(e + 1).Picture
Next e
End If
End Sub



Private Sub Form_Activate()
Unload Intro
Timer1.Enabled = True

End Sub

Private Sub Form_Initialize()
FormLoad
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Timer2.Enabled = True Then
Timer2.Enabled = False
Call Timer2_Timer
End If
If KeyCode = vbKeyEscape And Picture4.Enabled = True And Timer3.Enabled = False Then
Unload Me
End If
If KeyCode = vbKeyF1 And Dost <> "HELP" And Timer3.Enabled = False Then
BeginPlaySound 66
End If
If KeyCode = vbKeyF1 And UnDost = 0 And Dost = "HELP" Then
Call Cheat
End If
If KeyCode = vbKeyF10 And (Shift And vbShiftMask) Then
If Dost = "HELP" Then Exit Sub
Dost = "HELP"
BeginPlaySound 60
End If

End Sub

Private Sub FormLoad()

Set Pict = Picture4.Picture

Dost = ""
UnDost = 0
Num = Val(GetSetting("Puzzles", "Startup", "Level", "0"))
DoEvents
Label1.Left = (Picture4.Width - Label1.Width) \ 2
Intro.Label2.Visible = False
Picture4.Left = (Screen.Width \ 15 - Picture4.Width) \ 2
Picture4.Top = (Screen.Height \ 15 - Picture4.Height) \ 2
DoEvents
Intro.Label2.Caption = "Load data..."
Intro.Label2.Left = (Screen.Width - Intro.Label2.Width) \ 2
Intro.Label2.Top = Intro.Picture1.Top + Intro.Picture1.Height
Intro.Label2.Visible = True

Dim l, t
l = 80
t = 59
For e = 0 To 47
Load Picture2(e)
Picture2(e).Left = l
Picture2(e).Top = t
l = l + 80: If l > 640 Then l = 80: t = t + 80
Picture2(e).Visible = True
 Intro.Label2.Caption = "Load data " & e
 DoEvents
Next e

If Num > 0 Then
If Num <> 0 Then Num = Num - 1
Call NextLevel
Else
Call Puzzle
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If (MsgBox("ARE YOU SURE YOU WONT TO QUIT?", 4 + 48)) = vbYes Then
 EndPlaySound
 MsgBox "Visit me: http://www.i.com.ua/~aka"
 End
 Exit Sub
Else
 Cancel = True
End If

End Sub


Private Sub Picture2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> vbLeftButton Then Exit Sub
If MousePointer = 0 Then
 BeginPlaySound 62
 MouseIcon = Picture2(Index).Picture
 MousePointer = 99
 Picture2(Index) = Picture3.Picture
 i = Index
ElseIf MousePointer = 99 Then
 BeginPlaySound 61
 Picture2(i) = Picture2(Index)
 Picture2(Index) = MouseIcon
 MousePointer = 0
End If
End Sub

Private Sub Timer1_Timer()
For e = 0 To 47
 If Picture2(e) <> ImageList1.ListImages(e + 1).Picture Then Exit Sub
Next e
Picture4.Enabled = False
Timer2.Enabled = True
EndPlaySound
BeginPlaySound 64
Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()
Call NextLevel
Picture4.Enabled = True
Timer2.Enabled = False

End Sub

Private Sub Timer3_Timer()
Static tox
If tox = 0 Then
Picture4.BackColor = vbBlue
Picture4.PaintPicture Pict, 0, 0, , , , , , , vbSrcErase
tox = 1
ElseIf tox = 1 Then
Picture4.BackColor = vbRed
Picture4.PaintPicture Pict, 0, 0, , , , , , , vbSrcErase
tox = 2
ElseIf tox = 2 Then
Picture4.BackColor = vbMagenta
Picture4.PaintPicture Pict, 0, 0, , , , , , , vbSrcErase
tox = 3
ElseIf tox = 3 Then
Picture4.BackColor = vbCyan
Picture4.PaintPicture Pict, 0, 0, , , , , , , vbSrcErase
tox = 4
ElseIf tox = 4 Then
Picture4.BackColor = vbYellow
Picture4.PaintPicture Pict, 0, 0, , , , , , , vbSrcErase
tox = 0
Timer4.Enabled = True
End If
End Sub

Private Sub Timer4_Timer()
Timer4.Interval = 4000
Static Endd
Endd = Endd + 1
If Endd = 2 Then End
Picture1.Cls
Picture1.PaintPicture Intro.Picture1, 73, 52

End Sub
