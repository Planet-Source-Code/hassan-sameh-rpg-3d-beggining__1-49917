Attribute VB_Name = "Matrix"
'-------------------------------------------
'Camera
'-------------------------------------------
Public Function CameraMat()
D3DXMatrixIdentity matWorld
HassanD3DDevice.SetTransform D3DTS_WORLD, matWorld
HassanDIDevice.GetDeviceStateKeyboard HassanDIState
'Right
If HassanDIState.Key(205) <> 0 Then
Angle = Angle - TURN_SPEED
If Angle < 0 Then
Angle = Deg360 - (-Angle)
End If
End If
'Left
If HassanDIState.Key(203) <> 0 Then
Angle = Angle + TURN_SPEED
If Angle > Deg360 Then
Angle = 0 + (Angle - Deg360)
End If
End If
AngleConv = Deg360 - Angle
'Up
If HassanDIState.Key(200) <> 0 Then
    Translations
    If AllowMove = True Then
    CamX = CamX + (Sin(AngleConv) * WALK_SPEED)
    CamZ = CamZ + (Cos(AngleConv) * WALK_SPEED)
    End If
End If
'Down
If HassanDIState.Key(208) <> 0 Then
    CamX = CamX - (Sin(AngleConv) * (WALK_SPEED))
    CamZ = CamZ - (Cos(AngleConv) * (WALK_SPEED))
End If
If HassanDIState.Key(201) <> 0 Then
Pitch = Pitch + PITCH_SPEED
End If
If HassanDIState.Key(209) <> 0 Then
Pitch = Pitch - PITCH_SPEED
End If
D3DXMatrixIdentity matView
D3DXMatrixIdentity MatPos
D3DXMatrixIdentity MatRotation
D3DXMatrixIdentity MatLook
D3DXMatrixRotationY MatRotation, Angle
D3DXMatrixRotationX MatPitch, Pitch
D3DXMatrixMultiply MatLook, MatRotation, MatPitch
D3DXMatrixTranslation MatPos, -CamX, -CamY, -CamZ
D3DXMatrixMultiply matView, MatPos, MatLook
HassanD3DDevice.SetTransform D3DTS_VIEW, matView
DoEvents
D3DXMatrixPerspectiveFovLH matProj, Pi / 3, 1, 1, 1000000
HassanD3DDevice.SetTransform D3DTS_PROJECTION, matProj
End Function
'----------------------------------------------
'Moving
'----------------------------------------------
Public Function MoveMat(ByVal LocMovX As Long, ByVal LocMovY As Long, ByVal LocMovZ As Long)
D3DXMatrixIdentity matWorld
D3DXMatrixIdentity matMove
D3DXMatrixTranslation matMove, LocMovX, LocMovY, LocMovZ
D3DXMatrixMultiply matWorld, matWorld, matMove
D3DXMatrixScaling matMove, 1, 1, 1
D3DXMatrixMultiply matWorld, matWorld, matMove
HassanD3DDevice.SetTransform D3DTS_WORLD, matWorld
End Function
'----------------------------------------------
'RotationY
'----------------------------------------------
Public Function RotMatYStatic(speedY As Long)
D3DXMatrixRotationY matRot, Timer * speedY
HassanD3DDevice.SetTransform D3DTS_WORLD, matRot
End Function
'----------------------------------------------
'RotationX
'----------------------------------------------
Public Function RotMatXStatic(speedX As Long)
D3DXMatrixRotationX matRot, Timer * speedX
HassanD3DDevice.SetTransform D3DTS_WORLD, matRot
End Function
'----------------------------------------------
'RotationZ
'----------------------------------------------
Public Function RotMatZStatic(speedZ As Long)
D3DXMatrixRotationZ matRot, Timer * speedZ
HassanD3DDevice.SetTransform D3DTS_WORLD, matRot
End Function
