VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Discover Cairo"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox imgmenu 
      BackColor       =   &H80000012&
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5355
      ScaleWidth      =   4875
      TabIndex        =   3
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblDetailHigh 
      Alignment       =   2  'Center
      BackColor       =   &H80000015&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "High"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblDetailMedium 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Medium"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblDetailLow 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Low"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblunuse3 
      BackColor       =   &H80000012&
      Caption         =   "Details"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lbl32bit 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "32 - BIT"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lbl16bit 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "16 - BIT"
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblunuse2 
      BackColor       =   &H80000007&
      Caption         =   "Screen Depth"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblunuse1 
      BackColor       =   &H80000007&
      Caption         =   "Screen Size "
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblvirsion 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "V 1.0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lbl1024 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1024 - 768"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lbl800 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "800 - 600"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lbl600 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "640 - 480"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblexit 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lbloptions 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Options"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New Game"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
lbl600.BackColor = vbBlack
lbl800.BackColor = vbBlack
lbl1024.BackColor = vbRed
lbl16bit.BackColor = vbRed
'lblDetailLow.BackColor = vbBlack
'lblDetailMedium.BackColor = vbBlack
lblDetailHigh.BackColor = vbRed
End Sub
Private Sub lbl1024_Click()
lbl600.BackColor = vbBlack
lbl800.BackColor = vbBlack
lbl1024.BackColor = vbRed
End Sub
Private Sub lbl16bit_Click()
lbl16bit.BackColor = vbRed
End Sub
Private Sub lbl600_Click()
lbl600.BackColor = vbRed
lbl800.BackColor = vbBlack
lbl1024.BackColor = vbBlack
End Sub
Private Sub lbl800_Click()
lbl600.BackColor = vbBlack
lbl800.BackColor = vbRed
lbl1024.BackColor = vbBlack
End Sub
Private Sub lblDetailHigh_Click()
'lblDetailLow.BackColor = vbBlack
'lblDetailMedium.BackColor = vbBlack
'lblDetailHigh.BackColor = vbRed
End Sub
'Private Sub lblDetailLow_Click()
'lblDetailLow.BackColor = vbRed
'lblDetailMedium.BackColor = vbBlack
'lblDetailHigh.BackColor = vbBlack
'End Sub
'Private Sub lblDetailMedium_Click()
'lblDetailLow.BackColor = vbBlack
'lblDetailMedium.BackColor = vbRed
'lblDetailHigh.BackColor = vbBlack
'End Sub
Private Sub lblexit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblexit.BackColor = vbRed
End Sub
Private Sub lblexit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblexit.BackColor = vbBlack
End
End Sub
Private Sub lbloptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbloptions.BackColor = vbRed
End Sub
Private Sub lbloptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbloptions.BackColor = vbBlack
If imgmenu.Visible = True Then imgmenu.Visible = False Else imgmenu.Visible = True
End Sub
Private Sub lblStart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'lblStart.BackColor = vbRed
'If lblDetailLow.BackColor = vbRed Then GraphicsDetails = 1
'If lblDetailMedium.BackColor = vbRed Then GraphicsDetails = 2
If lblDetailHigh.BackColor = vbRed Then GraphicsDetails = 3
End Sub
Private Sub lblStart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStart.BackColor = vbBlack
Menu.Visible = False
Game.Visible = True
End Sub
