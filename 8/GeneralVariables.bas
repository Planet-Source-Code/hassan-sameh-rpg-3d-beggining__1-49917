Attribute VB_Name = "GeneralVariables"
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public HassanDX As New DirectX8
Public HassanD3D As Direct3D8
Public HassanD3DDevice As Direct3DDevice8
Public HassanD3DX As New D3DX8
Public HassanDI As DirectInput8
Public HassanDIDevice As DirectInputDevice8
Public HassanDIState As DIKEYBOARDSTATE
Public MainFont As D3DXFont
Public MainFontDesc As IFont
Public TextRect As RECT
Public Fnt As New StdFont
Public GraphicsDetails As Long
Public ResW, ResH As Integer: Public ResB As Long
Public Angle As Single
Public Pitch As Single
Public Const Pi As Single = 3.141592653
Public Const Deg90 As Single = Pi / 2
Public Const Deg180 As Single = Pi
Public Const Deg270 As Single = Pi * 1.5
Public Const Deg360 As Single = Pi * 2
Public Const TURN_SPEED = Deg90 / 20
Public matView As D3DMATRIX
Public MatRotation As D3DMATRIX
Public MatPitch As D3DMATRIX
Public MatLook As D3DMATRIX
Public MatPos As D3DMATRIX
Public matWorld As D3DMATRIX
Public matProj As D3DMATRIX
Public AngleConv As Single
Public CamX, CamY, CamZ As Long
Public bRunning As Boolean
Public WALK_SPEED As Double
Public PITCH_SPEED As Double
Public mm As New GenerateMesh
Public MousePositionX As Long
Public MousePositionY As Long
Public FPS_LastCheck As Long
Public FPS_Count As Long
Public FPS_Current As Integer
Public matMove As D3DMATRIX
Public matRot As D3DMATRIX
Public AllowMove As Boolean
Public LX As Long: Public LY As Long: Public LZ As Long

'-------------------------------------------------------
'Some Extra Variables Even I Will Try to make it as Array
'-------------------------------------------------------
Public HassanMesh0001 As D3DXMesh
Public HassanMesh0002 As D3DXMesh
Public HassanMesh0003 As D3DXMesh
Public HassanMesh0004 As D3DXMesh
Public HassanMesh0005 As D3DXMesh
Public HassanMesh0006 As D3DXMesh
Public HassanMesh0007 As D3DXMesh
Public HassanMesh0008 As D3DXMesh
Public HassanMesh0009 As D3DXMesh
Public HassanMesh00010 As D3DXMesh
Public HassanMesh00011 As D3DXMesh
Public HassanMesh00012 As D3DXMesh
Public HassanMesh00013 As D3DXMesh
Public HassanMesh00014 As D3DXMesh
Public HassanMesh00015 As D3DXMesh
Public HassanMesh00016 As D3DXMesh
Public HassanMesh00017 As D3DXMesh
Public HassanMesh00018 As D3DXMesh
Public HassanMesh00019 As D3DXMesh
Public HassanMesh00020 As D3DXMesh
Public HassanMesh00021 As D3DXMesh
Public HassanMesh00022 As D3DXMesh
Public HassanMesh00023 As D3DXMesh
Public HassanMesh00024 As D3DXMesh
Public HassanMesh00025 As D3DXMesh
Public HassanMesh00026 As D3DXMesh
Public HassanMesh00027 As D3DXMesh
Public HassanMesh00028 As D3DXMesh
Public HassanMesh00029 As D3DXMesh
Public HassanMesh00030 As D3DXMesh
Public HassanMesh00031 As D3DXMesh
Public HassanMesh00032 As D3DXMesh
Public HassanMesh00033 As D3DXMesh
Public HassanMesh00034 As D3DXMesh
Public HassanMesh00035 As D3DXMesh
Public HassanMesh00036 As D3DXMesh
Public HassanMesh00037 As D3DXMesh
Public HassanMesh00038 As D3DXMesh
Public HassanMesh00039 As D3DXMesh
Public HassanMesh00040 As D3DXMesh
Public HassanMesh00041 As D3DXMesh
Public HassanMesh00042 As D3DXMesh
Public HassanMesh00043 As D3DXMesh
Public HassanMesh00044 As D3DXMesh
Public HassanMesh00045 As D3DXMesh
Public HassanMesh00046 As D3DXMesh
Public HassanMesh00047 As D3DXMesh
Public HassanMesh00048 As D3DXMesh
Public HassanMesh00049 As D3DXMesh
Public HassanMesh00050 As D3DXMesh
Public HassanMesh00051 As D3DXMesh

Public HassanMeshTextures0001() As Direct3DTexture8
Public HassanMeshTextures0002() As Direct3DTexture8
Public HassanMeshTextures0003() As Direct3DTexture8
Public HassanMeshTextures0004() As Direct3DTexture8
Public HassanMeshTextures0005() As Direct3DTexture8
Public HassanMeshTextures0006() As Direct3DTexture8
Public HassanMeshTextures0007() As Direct3DTexture8
Public HassanMeshTextures0008() As Direct3DTexture8
Public HassanMeshTextures0009() As Direct3DTexture8
Public HassanMeshTextures00010() As Direct3DTexture8
Public HassanMeshTextures00011() As Direct3DTexture8
Public HassanMeshTextures00012() As Direct3DTexture8
Public HassanMeshTextures00013() As Direct3DTexture8
Public HassanMeshTextures00014() As Direct3DTexture8
Public HassanMeshTextures00015() As Direct3DTexture8
Public HassanMeshTextures00016() As Direct3DTexture8
Public HassanMeshTextures00017() As Direct3DTexture8
Public HassanMeshTextures00018() As Direct3DTexture8
Public HassanMeshTextures00019() As Direct3DTexture8
Public HassanMeshTextures00020() As Direct3DTexture8
Public HassanMeshTextures00021() As Direct3DTexture8
Public HassanMeshTextures00022() As Direct3DTexture8
Public HassanMeshTextures00023() As Direct3DTexture8
Public HassanMeshTextures00024() As Direct3DTexture8
Public HassanMeshTextures00025() As Direct3DTexture8
Public HassanMeshTextures00026() As Direct3DTexture8
Public HassanMeshTextures00027() As Direct3DTexture8
Public HassanMeshTextures00028() As Direct3DTexture8
Public HassanMeshTextures00029() As Direct3DTexture8
Public HassanMeshTextures00030() As Direct3DTexture8
Public HassanMeshTextures00031() As Direct3DTexture8
Public HassanMeshTextures00032() As Direct3DTexture8
Public HassanMeshTextures00033() As Direct3DTexture8
Public HassanMeshTextures00034() As Direct3DTexture8
Public HassanMeshTextures00035() As Direct3DTexture8
Public HassanMeshTextures00036() As Direct3DTexture8
Public HassanMeshTextures00037() As Direct3DTexture8
Public HassanMeshTextures00038() As Direct3DTexture8
Public HassanMeshTextures00039() As Direct3DTexture8
Public HassanMeshTextures00040() As Direct3DTexture8
Public HassanMeshTextures00041() As Direct3DTexture8
Public HassanMeshTextures00042() As Direct3DTexture8
Public HassanMeshTextures00043() As Direct3DTexture8
Public HassanMeshTextures00044() As Direct3DTexture8
Public HassanMeshTextures00045() As Direct3DTexture8
Public HassanMeshTextures00046() As Direct3DTexture8
Public HassanMeshTextures00047() As Direct3DTexture8
Public HassanMeshTextures00048() As Direct3DTexture8
Public HassanMeshTextures00049() As Direct3DTexture8
Public HassanMeshTextures00050() As Direct3DTexture8
Public HassanMeshTextures00051() As Direct3DTexture8
