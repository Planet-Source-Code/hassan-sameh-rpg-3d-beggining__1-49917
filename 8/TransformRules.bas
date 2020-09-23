Attribute VB_Name = "TransformRules"
'it's supposed to be the collision detiction but i'm not satisfyed on the way it's is
' i will try to find a better way for it
Public Function Translations()
Call mm.TransRulesForward(-20, 70, -100, 100): If AllowMove = False Then GoTo GetOut
Call mm.TransRulesForward(-480, -470, -470, 470): If AllowMove = False Then GoTo GetOut
Call mm.TransRulesForward(-470, 470, 470, 480): If AllowMove = False Then GoTo GetOut
Call mm.TransRulesForward(470, 480, -470, 470): If AllowMove = False Then GoTo GetOut
Call mm.TransRulesForward(-470, 470, -480, -470): If AllowMove = False Then GoTo GetOut
GetOut:
Exit Function
End Function
