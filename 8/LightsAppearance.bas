Attribute VB_Name = "LightsAppearance"
Public Sub SetupLights()
Dim COL As D3DCOLORVALUE
Dim MTRL As D3DMATERIAL8

With COL: .r = 1: .g = 1: .B = 1: .a = 1: End With
'r = red : b = green : g = blue : a = alpha

MTRL.diffuse = COL
MTRL.Ambient = COL

HassanD3DDevice.SetMaterial MTRL

Dim light As D3DLIGHT8
light.Type = D3DLIGHT_POINT
light.Position.X = 0
light.Position.Y = 300
light.Position.z = -600

light.diffuse.B = 0.1
light.diffuse.g = 0.6
light.diffuse.r = 0.4

light.Range = 5000


HassanD3DDevice.SetLight 0, light
HassanD3DDevice.LightEnable 0, 1
HassanD3DDevice.SetRenderState D3DRS_LIGHTING, 1
End Sub
Public Sub SetupLights2()
Dim COL2 As D3DCOLORVALUE
Dim MTRL2 As D3DMATERIAL8

With COL2: .r = 1: .g = 1: .B = 1: .a = 1: End With
'r = red : b = green : g = blue : a = alpha

MTRL2.diffuse = COL2
MTRL2.Ambient = COL2

HassanD3DDevice.SetMaterial MTRL2

Dim light2 As D3DLIGHT8
light2.Type = D3DLIGHT_POINT
light2.Position.X = 0
light2.Position.Y = 300
light2.Position.z = 600

light2.diffuse.B = 0.1
light2.diffuse.g = 0.6
light2.diffuse.r = 0.4

light2.Range = 5000


HassanD3DDevice.SetLight 0, light2
HassanD3DDevice.LightEnable 0, 1
HassanD3DDevice.SetRenderState D3DRS_LIGHTING, 1
End Sub
