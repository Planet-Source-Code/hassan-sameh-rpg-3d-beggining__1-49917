VERSION 5.00
Begin VB.Form Game 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   5550
   ClientTop       =   4155
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   ScaleHeight     =   136
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim B As Boolean
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyZ: CamY = CamY + 10
Case vbKeyX: CamY = CamY - 10
Case vbKeyL: LX = LX + 3
Case vbKeyJ: LX = LX - 3
Case vbKeyI: LZ = LZ + 3
Case vbKeyK: LZ = LZ - 3
End Select

End Sub
Private Sub Form_Load()
'Show The Form
Me.Show
'The Beggining Values For The Variable
B = InitValues()
'Retrive Which Screen Size
B = RetriveResolution()
'Initialize The Direct Input
B = InitDI()
'Initialize The Direct3D
B = InitD3D()
'Drow Loading
B = InitLoad()
'Load The X Files And Textures
B = InitSceneMeshsAndTextures()
'Place The Mouse Pointer Position
B = SetCursorPos(500, 350)
'Hide The Mouse Pointer
B = ShowCursor(0)
'Render
While bRunning
RENDER
If GetTickCount() - FPS_LastCheck >= 100 Then
FPS_Current = FPS_Count * 10
FPS_Count = 0
FPS_LastCheck = GetTickCount()
End If
FPS_Count = FPS_Count + 1
DoEvents
Wend
End Sub
Private Function RetriveResolution() As Boolean
If Menu.lbl600.BackColor = vbRed Then ResW = 640: ResH = 480
If Menu.lbl800.BackColor = vbRed Then ResW = 800: ResH = 600
If Menu.lbl1024.BackColor = vbRed Then ResW = 1024: ResH = 768
End Function
Private Function InitDI() As Boolean
'Devine The Variable For Direct Input
Set HassanDI = HassanDX.DirectInputCreate()
'Retrive Errors
If Err.Number <> 0 Then MsgBox "Error starting Direct Input, please make sure you have DirectX installed", vbApplicationModal: End
'Create The Direct Input Device
Set HassanDIDevice = HassanDI.CreateDevice("GUID_SysKeyboard")
HassanDIDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
HassanDIDevice.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
HassanDIDevice.Acquire
End Function
Private Function InitD3D() As Boolean
On Local Error Resume Next
Set HassanD3D = HassanDX.Direct3DCreate()
If HassanD3D Is Nothing Then Exit Function
'Variable Hold The ScreenMode
Dim Mode As D3DDISPLAYMODE
'Retrive The Screen Mode
HassanD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode
'Defie The Present_Parameters
Dim D3DPP As D3DPRESENT_PARAMETERS
'If HassanD3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DFMT_D16, D3DFMT_D16, False) Then MsgBox "it's Working"
D3DPP.SwapEffect = D3DSWAPEFFECT_FLIP
D3DPP.BackBufferFormat = Mode.Format
D3DPP.BackBufferCount = 1
D3DPP.EnableAutoDepthStencil = 1
D3DPP.AutoDepthStencilFormat = D3DFMT_D16
D3DPP.BackBufferHeight = ResH
D3DPP.BackBufferWidth = ResW
D3DPP.hDeviceWindow = Me.hWnd
'Create Direct3D Device
Set HassanD3DDevice = HassanD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)
If HassanD3DDevice Is Nothing Then Exit Function
'No Cull
HassanD3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
HassanD3DDevice.SetRenderState D3DRS_ZENABLE, 1
'No Light
HassanD3DDevice.SetRenderState D3DRS_LIGHTING, 0
InitD3D = True
End Function
Private Function InitLoad() As Boolean
Fnt.Name = "Verdana"
Fnt.Size = 8
Fnt.Bold = True
Set MainFontDesc = Fnt
Set MainFont = HassanD3DX.CreateFont(HassanD3DDevice, MainFontDesc.hFont)
DrawLoad ("")
End Function
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
B = ShowCursor(0)
End
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > MousePositionX Then
Angle = Angle - TURN_SPEED
If Angle < 0 Then
Angle = Deg360 - (-Angle)
End If
End If
If X < MousePositionX Then
Angle = Angle + TURN_SPEED
If Angle > Deg360 Then
Angle = 0 + (Angle - Deg360)
End If
End If
If Y > MousePositionY Then
Pitch = Pitch - PITCH_SPEED
End If
If Y < MousePositionY Then
Pitch = Pitch + PITCH_SPEED
End If
MousePositionX = X
MousePositionY = Y
If X > 900 Then B = SetCursorPos(500, 350)
If X < 100 Then B = SetCursorPos(500, 350)
If Y > 700 Then B = SetCursorPos(500, 350)
If Y < 100 Then B = SetCursorPos(500, 350)
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set HassanDX = Nothing
Set HassanD3D = Nothing
Set HassanD3DDevice = Nothing
Set HassanD3DX = Nothing
Set HassanDI = Nothing
Set HassanDIDevice = Nothing
Set HassanMesh0001 = Nothing
Set HassanMesh0002 = Nothing
Set HassanMeshTextures0001 = Nothing
Set HassanMeshTextures0002 = Nothing
B = ShowCursor(1)
End Sub
Private Function InitValues()
CamX = 0: CamY = 45: CamZ = -130
bRunning = True
WALK_SPEED = 4
PITCH_SPEED = 0.025
TextRect.Top = 5
TextRect.Left = 100
TextRect.bottom = 500
TextRect.Right = 1000
LX = 0: LY = 20: LZ = 0
End Function
