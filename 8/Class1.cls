VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenerateMesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------
'                  Next..!!
'-------------------------------------------------
'1-Collision..( Not Completed..!! )
'2-Character Animation..( Not Yet  )
'3-Mouse Pointer Stuff..( Not Complited  )
'4-InitD3D Stuff..( Not Yet )

'--------------------------------------------------
'Load The X File
'--------------------------------------------------
Public Function AddMeshStatic(XName As D3DXMesh, TextureName() As Direct3DTexture8, XPath As String, TexturePath As String)
On Local Error Resume Next
Dim MtrlBuffer As D3DXBuffer
Dim I As Long
Dim HassanNumMaterials As Long
Set XName = HassanD3DX.LoadMeshFromX(App.Path + XPath, 544, HassanD3DDevice, Nothing, MtrlBuffer, HassanNumMaterials)
If XName Is Nothing Then MsgBox "Can't find " + XPath + " This Application will be turminated": End
ReDim TextureName(HassanNumMaterials)
Set TextureName(I) = HassanD3DX.CreateTextureFromFile(HassanD3DDevice, App.Path + TexturePath)
If TextureName(I) Is Nothing Then MsgBox "Can't find " + TexturePath + " This Application will be turminated": End
Set MtrlBuffer = Nothing
DrawLoad ("loading...." + XPath + "      " + TexturePath)
End Function
'-----------------------------------------------------
'Rendering
'-----------------------------------------------------
Public Function RenderStatic(HasTexture() As Direct3DTexture8, HasX As D3DXMesh, RotateYstatic As Boolean, RotateXstatic As Boolean, RotateZstatic As Boolean, ByVal SpeedRotate As Long)
If RotateYstatic = True Then RotMatYStatic (SpeedRotate)
If RotateXstatic = True Then RotMatXStatic (SpeedRotate)
If RotateZstatic = True Then RotMatZStatic (SpeedRotate)
HassanD3DDevice.BeginScene
HassanD3DDevice.SetTexture 0, HasTexture(I)
HasX.DrawSubset I
HassanD3DDevice.EndScene
TextRect.Top = 5
TextRect.Left = 1
TextRect.bottom = 480
TextRect.Right = 640
HassanD3DX.DrawText MainFont, &HFFCCCCFF, "Current Frame Rate(CFR) : " & FPS_Current, TextRect, DT_TOP Or DT_CENTER
TextRect.Top = 15
TextRect.Left = 50
TextRect.bottom = 480
TextRect.Right = 640
HassanD3DX.DrawText MainFont, &HFFCCCCFF, "CameraX: " & CamX & " CameraY:" & CamY & " CameraZ:" & CamZ & "                       " & Angle, TextRect, DT_TOP Or DT_CENTER
End Function
'-------------------------------------------------------------------------------
'That is for the collision Detection But i hope to improve it in the future
'--------------------------------------------------------------------------------
Public Function TransRulesForward(X1Rule As Long, X2Rule As Long, Z1Rule As Long, Z2Rule As Long)
If CamX >= X1Rule And CamX <= X2Rule And CamZ >= Z1Rule And CamZ <= Z2Rule Then
AllowMove = False
    If Angle >= 0 And Angle <= 1.5 Then
    CamX = CamX - WALK_SPEED
    ElseIf Angle > 4.7 And Angle < 7.6 Then
    CamX = CamX + WALK_SPEED
    End If
Else
AllowMove = True
End If
End Function
