VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mMesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_X As Single
Private m_Y As Single
Private m_Z As Single
Private m_xTheta As Single
Private m_yTheta As Single
Private m_zTheta As Single
Private m_Color(2) As Long
Private m_Vectcount As Integer
Private m_Vector() As Single
Private m_Normcount As Integer
Private m_Normal() As Integer
Private m_Zorder() As Integer
'yea...

Public Property Let x(ByVal vData As Single)
    m_X = vData
End Property

Public Property Get x() As Single
    x = m_X
End Property

Public Property Let y(ByVal vData As Single)
    m_Y = vData
End Property

Public Property Get y() As Single
    y = m_Y
End Property

Public Property Let z(ByVal vData As Single)
    m_Z = vData
End Property

Public Property Get z() As Single
    z = m_Z
End Property

Public Property Let xTheta(ByVal vData As Single)
    If vData > 360 Then
        vData = 0
    End If
    If vData < 0 Then
        vData = 360
    End If
    m_xTheta = vData
End Property

Public Property Get xTheta() As Single
    xTheta = m_xTheta
End Property

Public Property Let yTheta(ByVal vData As Single)
    If vData > 360 Then
        vData = 0
    End If
    If vData < 0 Then
        vData = 360
    End If
    m_yTheta = vData
End Property

Public Property Get yTheta() As Single
    yTheta = m_yTheta
End Property

Public Property Let zTheta(ByVal vData As Single)
    If vData > 360 Then
        vData = 0
    End If
    If vData < 0 Then
        vData = 360
    End If
    m_zTheta = vData
End Property

Public Property Get zTheta() As Single
    zTheta = m_zTheta
End Property

Public Property Let Color(index As Integer, ByVal vData As Long)
    m_Color(index) = vData
End Property

Public Property Get Color(index As Integer) As Long
    Color = m_Color(index)
End Property

Public Property Let VectCount(ByVal vData As Long)
    m_Vectcount = vData - 1
    ReDim m_Vector(0 To (m_Vectcount), 2, 2)
End Property

Public Property Get VectCount() As Long
    VectCount = m_Vectcount + 1
End Property


Public Property Let normcount(ByVal vData As Long)
    m_Normcount = vData - 1
    ReDim m_Normal(0 To (m_Normcount), 2, 20)
    ReDim m_Zorder(0 To (m_Normcount))
End Property

Public Property Get normcount() As Long
    normcount = m_Normcount + 1
End Property

Public Property Let Vector(index As Integer, real As Integer, xyz As Integer, ByVal vData As Single)
    m_Vector(index, real, xyz) = vData
End Property

Public Property Get Vector(index As Integer, real As Integer, xyz As Integer) As Single
On Error Resume Next
    Vector = m_Vector(index, real, xyz)
End Property

Public Property Let Normal(index As Integer, prop As Integer, vertex As Integer, ByVal vData As Long)
    m_Normal(index, prop, vertex) = vData
End Property

Public Property Get Normal(index As Integer, prop As Integer, vertex As Integer) As Long
    Normal = m_Normal(index, prop, vertex)
End Property

Public Property Let Zorder(index As Integer, ByVal vData As Long)
    m_Zorder(index) = vData
End Property

Public Property Get Zorder(index As Integer) As Long
    Zorder = m_Zorder(index)
End Property

Public Function Render(hdest As Object)


    Dim i As Integer
    Dim j As Integer
    Dim temp1
    Dim BackHDC As Long
    BackHDC = hdest.hdc
    
    i = 0

    Randomize
    For i = 0 To m_Normcount
        Dim colr
            j = 0
        Do While m_Normal(i, 0, j) >= 0
            Dim nx As Single, ny As Single, nz As Single
            Dim ox As Single, oy As Single, oz As Single
            'do the x axis first
            ox = Vector(m_Normal(i, 0, j), 0, 0)
            oy = Vector(m_Normal(i, 0, j), 0, 1)
            oz = Vector(m_Normal(i, 0, j), 0, 2)
            nx = ox
            ny = Cos(RadAng(m_xTheta)) * oy - Sin(RadAng(m_xTheta)) * oz
            nz = Cos(RadAng(m_xTheta)) * oz + Sin(RadAng(m_xTheta)) * oy
            'then the y
            ox = nx
            oy = ny
            oz = nz
            nx = Cos(RadAng(m_yTheta)) * ox - Sin(RadAng(m_yTheta)) * oz
            ny = oy
            nz = Cos(RadAng(m_yTheta)) * oz + Sin(RadAng(m_yTheta)) * ox
            'then the z
            ox = nx
            oy = ny
            oz = nz
            nx = Cos(RadAng(m_zTheta)) * ox - Sin(RadAng(m_zTheta)) * oy
            ny = Cos(RadAng(m_zTheta)) * oy + Sin(RadAng(m_zTheta)) * ox
            nz = oz
            
            nx = nx - m_X
            ny = ny - m_Y
            nz = nz - m_Z
            
            colr = colr + nz
            Vector(m_Normal(i, 0, j), 1, 0) = (256 * nx / (nz) + hdest.ScaleWidth / 2) * 0.00000024 + 2  '* 0.00001 + 2
            Vector(m_Normal(i, 0, j), 1, 1) = (256 * ny / (nz) + hdest.ScaleHeight / 2) * 0.00000024 + 2  '* 0.00001 + 2
           
            j = j + 1
            m_Normal(i, 2, 0) = j
           DoEvents
        Loop
        colr = colr / j
        
        m_Normal(i, 1, 0) = colr
        
colr = 100 + colr
Next i
Dim k As Integer
Dim o As Integer
Dim L
For L = 0 To m_Normcount
    colr = 100 + m_Normal((L), 1, 0)
    ReDim nPoint(m_Normal((L), 2, 0) - 1) As POINTAPI
    Dim m
    For m = 0 To m_Normal((L), 2, 0) - 1
        nPoint(m).x = Vector(m_Normal((L), 0, m), 1, 0)
        nPoint(m).y = Vector(m_Normal((L), 0, m), 1, 1)
    Next m
    colr = Abs((Abs(colr) / 50) - 255) - 127
    hdest.ForeColor = (RGB(colr + m_Color(0), colr + m_Color(1), colr + m_Color(2)))
    hdest.FillColor = (RGB(colr + m_Color(0), colr + m_Color(1), colr + m_Color(2)))
    hdest.FillStyle = vbFSSolid
    drPolygon hdest.hdc, nPoint(0), m_Normal(L, 2, 0)
    hdest.ForeColor = vbRed

Next L
 '   TextOut hdest.hdc, 10, 10, ("x-axis: " & m_xTheta), 11
 '   TextOut hdest.hdc, 10, 26, ("y-axis: " & m_yTheta), 11
 '   TextOut hdest.hdc, 10, 42, ("z-axis: " & m_zTheta), 11
End Function

Private Sub Class_Initialize()
m_X = 0
m_Y = 0
m_Z = 0
m_xTheta = 0
m_yTheta = 0
m_zTheta = 0
End Sub
