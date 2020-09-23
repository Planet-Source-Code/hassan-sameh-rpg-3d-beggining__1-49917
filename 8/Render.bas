Attribute VB_Name = "RenderScene"
Public Sub RENDER()
HassanD3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 88799, 1#, 0
'---------------------------------------------------------
'Load Static Mechs
'---------------------------------------------------------
B = mm.RenderStatic(HassanMeshTextures0001(), HassanMesh0001, False, False, False, 0)
B = mm.RenderStatic(HassanMeshTextures0003(), HassanMesh0004, False, False, False, 0)
B = mm.RenderStatic(HassanMeshTextures0004(), HassanMesh0009, False, False, False, 0)
B = mm.RenderStatic(HassanMeshTextures0004(), HassanMesh00010, False, False, False, 0)
B = mm.RenderStatic(HassanMeshTextures0004(), HassanMesh00011, False, False, False, 0)
B = mm.RenderStatic(HassanMeshTextures0004(), HassanMesh00012, False, False, False, 0)
B = mm.RenderStatic(HassanMeshTextures0003(), HassanMesh00013, False, False, False, 0)
'----------------------------------------------------------
'Load meshs with matrix Rotations
'----------------------------------------------------------
B = mm.RenderStatic(HassanMeshTextures0003(), HassanMesh00015, True, False, False, 4)
CameraMat
'B = SetupLights(LX, LY, LZ)
HassanD3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub
'------------------------------------------------------------
'Render Loading
'------------------------------------------------------------
Public Function DrawLoad(TxtDraw As String)
HassanD3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 1, 1#, 0
HassanD3DX.DrawText MainFont, &HFFCCCCFF, TxtDraw, TextRect, DT_TOP Or DT_CENTER
HassanD3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Function
