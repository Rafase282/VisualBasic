Attribute VB_Name = "Module1"
Public Angle As Single
Public Const rd = 3.14159 / 180

' ----------------------
' *���w���ת����u*
' ----------------------
'
Public Sub Movel(Leng As Single, Col As Long)
    x = Leng * Cos(rd * Angle)
    y = Leng * Sin(rd * Angle)
    Form1.Line -Step(x, -y), Col
End Sub

' -------------------------
' *�]�w��V��*
' -------------------------
'
Public Sub SetAngle(a As Single)
    Angle = a
End Sub

' -------------------------
' *�]�w�{�b��m*
' -------------------------
'
Public Sub SetPoint(x As Single, y As Single)
    Form1.CurrentX = x
    Form1.CurrentY = y
End Sub

' -------------------------
' *��V*
' -------------------------
'
Public Sub Turn(a As Single)
    Angle = Angle + a
    Angle = Angle - Int(Angle) + Int(Angle) Mod 360
End Sub