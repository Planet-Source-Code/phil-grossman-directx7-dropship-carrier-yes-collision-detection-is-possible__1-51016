VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Run Game"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2115
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   103.773
   ScaleMode       =   0  'User
   ScaleWidth      =   160.227
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   18435
      Left            =   2460
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   200
      TabIndex        =   5
      Top             =   495
      Width           =   7410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run Game"
      Height          =   600
      Left            =   345
      TabIndex        =   4
      Top             =   75
      Width           =   1575
   End
   Begin VB.PictureBox LoadPB 
      AutoRedraw      =   -1  'True
      Height          =   1572
      Left            =   4500
      ScaleHeight     =   1515
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   4545
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.PictureBox SavePB 
      AutoRedraw      =   -1  'True
      Height          =   1572
      Left            =   6840
      ScaleHeight     =   1515
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   4575
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   465
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1095
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.FileListBox File1 
      Height          =   3600
      Left            =   420
      Pattern         =   "*.jpg*"
      TabIndex        =   0
      Top             =   1755
      Visible         =   0   'False
      Width           =   3420
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show

End Sub

' convert jpgs into bitmaps
' if they're already converted, run the game
Private Sub Form_Load()

Dim mypic, mysavepb
File1.Path = App.Path & "\textures\"
'If File1.ListCount = 7 Then ' adjust this if you add more textures
    For Genc = 0 To File1.ListCount - 1
        mysavepb = Left(File1.List(Genc), 10)
        Text1.Text = File1.ListCount & " " & File1.List(Genc) & " " & mysavepb
        mypic = File1.Path & "\" & File1.List(Genc)
        LoadPB = LoadPicture(mypic)
        SavePB = LoadPB
        SavePicture SavePB.Picture, App.Path & "\textures\" & mysavepb & ".bmp"
    Next Genc
    Picture1 = LoadPicture(App.Path & "\textures\" & "PLAN000001.bmp")
Picture1.ScaleHeight = 300
Picture1.ScaleWidth = 600
'End If

End Sub

