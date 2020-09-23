Attribute VB_Name = "LightStuff"
Public Function SetupLights(LightX As Long, LightY As Long, LightZ As Long)
Dim COL As D3DCOLORVALUE
Dim MTRL As D3DMATERIAL8
With COL: .r = 1: .g = 1: .B = 1: .a = 1: End With
'r = red : b = green : g = blue : a = alpha
MTRL.diffuse = COL
MTRL.Ambient = COL
HassanD3DDevice.SetMaterial MTRL
Dim light As D3DLIGHT8
light.Type = D3DLIGHT_POINT
light.Position.X = LightX
light.Position.Y = LightY
light.Position.z = LightZ
light.diffuse.B = 1
light.diffuse.g = 1
light.diffuse.r = 1
light.Range = 20000
HassanD3DDevice.SetLight 0, light
HassanD3DDevice.LightEnable 0, 1
HassanD3DDevice.SetRenderState D3DRS_LIGHTING, 1
End Function
