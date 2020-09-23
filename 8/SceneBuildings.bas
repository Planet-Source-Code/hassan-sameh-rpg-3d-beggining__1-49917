Attribute VB_Name = "SceneBuildings"
Public Function InitSceneMeshsAndTextures()
'-------------------------------------------
              'low Details
'-------------------------------------------
'If GraphicsDetails = 1 Then
'B = mm.AddMeshStatic(HassanMesh0001, HassanMeshTextures0001(), "\data\3d\3d0001.x", "\data\bitmap\texture0001.jpg")
'End If
'-------------------------------------------
              'Medium Details
'-------------------------------------------
'If GraphicsDetails = 2 Then
'B = mm.AddMeshStatic(HassanMesh0001, HassanMeshTextures0001(), "\data\3d\3d0001.x", "\data\bitmap\texture0001.jpg")
'End If
'-------------------------------------------
               'high Details
'-------------------------------------------
If GraphicsDetails = 3 Then
B = mm.AddMeshStatic(HassanMesh0001, HassanMeshTextures0001(), "\data\3d\3d0001.x", "\data\bitmap\texture0003.jpg")
B = mm.AddMeshStatic(HassanMesh0004, HassanMeshTextures0003(), "\data\3d\3d0004.x", "\data\bitmap\texture0003.jpg")
B = mm.AddMeshStatic(HassanMesh0009, HassanMeshTextures0004(), "\data\3d\3d0009.x", "\data\bitmap\texture0003.jpg")
B = mm.AddMeshStatic(HassanMesh00010, HassanMeshTextures0004(), "\data\3d\3d00010.x", "\data\bitmap\texture0003.jpg")
B = mm.AddMeshStatic(HassanMesh00011, HassanMeshTextures0004(), "\data\3d\3d00011.x", "\data\bitmap\texture0003.jpg")
B = mm.AddMeshStatic(HassanMesh00012, HassanMeshTextures0004(), "\data\3d\3d00012.x", "\data\bitmap\texture0003.jpg")
B = mm.AddMeshStatic(HassanMesh00013, HassanMeshTextures0003(), "\data\3d\3d00013.x", "\data\bitmap\texture0003.jpg")
B = mm.AddMeshStatic(HassanMesh00015, HassanMeshTextures0003(), "\data\3d\3d00015.x", "\data\bitmap\texture0003.jpg")
End If
End Function
