Attribute VB_Name = "THREEDENGINE"
Type POINTAPI
x As Single
y As Single
End Type
Public Type Vector
x As Long
y As Long
z As Long
End Type
Public Type Normal
verts() As Integer
End Type
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function drPolygon Lib "gdi32" Alias "Polygon" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Const PI = 3.14159265358979
Public Function RadAng(Radian As Single) As Single
On Error Resume Next
RadAng = (PI * Radian) / 180
End Function
