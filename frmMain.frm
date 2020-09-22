VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3d Tunnel"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   474
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox prim 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   7065
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.Timer tmr1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   360
         Top             =   3000
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Game"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tunnel Spoof in 3d! v1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Index           =   1
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tunnel Spoof in 3d! v1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   0
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Caption         =   "Made by Daneish (lol =P)"
         Height          =   615
         Left            =   3000
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim walls(20) As New mMesh
Dim nl
Dim nv
Dim cWidth As Single
Dim player As New mMesh
Dim up, le, dw, ri
Dim running
Dim wall As Single
Dim ceil As Single
Private Sub main_loop()
Do

If running = False Then
Exit Do

End If
DoEvents
Loop


End Sub
Private Sub cmdStart_Click()
cmdStart.visible = False
lblTitle(0).visible = False
lblTitle(1).visible = False
lblName.visible = False
running = True
tmr1.Enabled = True
Me.WindowState = 2
Me.BorderStyle = 0
'main_loop
End Sub

Private Sub Form_Load()
Dim i
Me.WindowState = 0
Me.BorderStyle = 1
tmr1.Enabled = False
running = False
up = False
dw = False
le = False
ri = False
For i = 20 To 0 Step -1
With walls(i)
.normcount = 4
.VectCount = 8
Dim cs As Long

cs = 150
'vertex placement for cube
.Vector(0, 0, 0) = -cs   'upper front left = 0
.Vector(0, 0, 1) = cs
.Vector(0, 0, 2) = cs / 2
.Vector(1, 0, 0) = cs    'upper front right = 1
.Vector(1, 0, 1) = cs
.Vector(1, 0, 2) = cs / 2
.Vector(2, 0, 0) = -cs   'upper back left = 2
.Vector(2, 0, 1) = cs
.Vector(2, 0, 2) = -cs / 2
.Vector(3, 0, 0) = cs    'upper back right = 3
.Vector(3, 0, 1) = cs
.Vector(3, 0, 2) = -cs / 2
.Vector(4, 0, 0) = -cs   'lower front left = 4
.Vector(4, 0, 1) = -cs
.Vector(4, 0, 2) = cs / 2
.Vector(5, 0, 0) = cs    'lower front right = 5
.Vector(5, 0, 1) = -cs
.Vector(5, 0, 2) = cs / 2
.Vector(6, 0, 0) = -cs   'lower back left = 6
.Vector(6, 0, 1) = -cs
.Vector(6, 0, 2) = -cs / 2
.Vector(7, 0, 0) = cs    'lower back right = 7
.Vector(7, 0, 1) = -cs
.Vector(7, 0, 2) = -cs / 2
'normal definition for cube
.Normal(0, 0, 0) = 2      'top normal
.Normal(0, 0, 1) = 3
.Normal(0, 0, 2) = 1
.Normal(0, 0, 3) = 0
.Normal(0, 0, 4) = -1
.Normal(1, 0, 0) = 4      'bottom normal
.Normal(1, 0, 1) = 5
.Normal(1, 0, 2) = 7
.Normal(1, 0, 3) = 6
.Normal(1, 0, 4) = -1
.Normal(2, 0, 0) = 0      'left normal
.Normal(2, 0, 1) = 4
.Normal(2, 0, 2) = 6
.Normal(2, 0, 3) = 2
.Normal(2, 0, 4) = -1
.Normal(3, 0, 0) = 1      'right normal
.Normal(3, 0, 1) = 5
.Normal(3, 0, 2) = 7
.Normal(3, 0, 3) = 3
.Normal(3, 0, 4) = -1
.Color(0) = 100
.Color(1) = 0
.Color(2) = 0
.z = Abs((i * (cs))) + 76  ' - (cs * 33))
End With

With player
.VectCount = 8
.normcount = 6
cs = 10
'vertex placement for cube
.Vector(0, 0, 0) = 0   'upper front left = 0
.Vector(0, 0, 1) = 0
.Vector(0, 0, 2) = -cs
.Vector(1, 0, 0) = 0    'upper front right = 1
.Vector(1, 0, 1) = 0
.Vector(1, 0, 2) = -cs
.Vector(2, 0, 0) = -cs * 1.5 'upper back left = 2
.Vector(2, 0, 1) = cs / 2
.Vector(2, 0, 2) = cs * 2
.Vector(3, 0, 0) = cs * 1.5  'upper back right = 3
.Vector(3, 0, 1) = cs / 2
.Vector(3, 0, 2) = cs * 2
.Vector(4, 0, 0) = 0   'lower front left = 4
.Vector(4, 0, 1) = 0
.Vector(4, 0, 2) = -cs
.Vector(5, 0, 0) = 0    'lower front right = 5
.Vector(5, 0, 1) = 0
.Vector(5, 0, 2) = -cs
.Vector(6, 0, 0) = 0 'lower back left = 6
.Vector(6, 0, 1) = -cs
.Vector(6, 0, 2) = cs
.Vector(7, 0, 0) = 0    'lower back right = 7
.Vector(7, 0, 1) = -cs
.Vector(7, 0, 2) = cs
'normal definition for cube
.Normal(0, 0, 0) = 2      'top normal
.Normal(0, 0, 1) = 3
.Normal(0, 0, 2) = 1
.Normal(0, 0, 3) = 0
.Normal(0, 0, 4) = -1
.Normal(1, 0, 0) = 4      'bottom normal
.Normal(1, 0, 1) = 5
.Normal(1, 0, 2) = 7
.Normal(1, 0, 3) = 6
.Normal(1, 0, 4) = -1
.Normal(2, 0, 0) = 0      'front normal
.Normal(2, 0, 1) = 1
.Normal(2, 0, 2) = 5
.Normal(2, 0, 3) = 4
.Normal(2, 0, 4) = -1
.Normal(3, 0, 0) = 2      'back normal
.Normal(3, 0, 1) = 3
.Normal(3, 0, 2) = 7
.Normal(3, 0, 3) = 6
.Normal(3, 0, 4) = -1
.Normal(4, 0, 0) = 0      'left normal
.Normal(4, 0, 1) = 4
.Normal(4, 0, 2) = 6
.Normal(4, 0, 3) = 2
.Normal(4, 0, 4) = -1
.Normal(5, 0, 0) = 1      'right normal
.Normal(5, 0, 1) = 5
.Normal(5, 0, 2) = 7
.Normal(5, 0, 3) = 3
.Normal(5, 0, 4) = -1
.z = 150
.y = 10
.x = 0
.Color(0) = 0
.Color(1) = 0
.Color(2) = 200
End With

Next i
nl = 0
nv = 0
cWidth = 300

wall = 298
ceil = 298
End Sub

Private Sub form_Resize()
prim.Width = Me.ScaleWidth
prim.Height = Me.ScaleHeight
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = vbKeyUp
up = True
Case Is = vbKeyDown
dw = True
Case Is = vbKeyLeft
le = True
Case Is = vbKeyRight
ri = True
Case Is = vbKeyReturn
cmdStart_Click
End Select
End Sub

Private Sub form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = vbKeyUp
up = False
Case Is = vbKeyDown
dw = False
Case Is = vbKeyLeft
le = False
Case Is = vbKeyRight
ri = False
End Select
End Sub

Private Sub tmr1_Timer()
Dim i
Dim max
max = 75

nl = nl + ((Rnd * max) - max / 2)
If nl > (wall - (cWidth / 2)) Then
nl = (wall - (cWidth / 2))
End If
If nl < (-wall + (cWidth / 2)) Then
nl = (-wall + (cWidth / 2))
End If

nv = nv + ((Rnd * max) - max / 2)
If nv > (ceil - (cWidth / 2)) Then
nv = (ceil - (cWidth / 2))
End If
If nv < (-ceil + (cWidth / 2)) Then
nv = (-ceil + (cWidth / 2))
End If
player.Color(0) = Rnd * 255
player.Color(1) = Rnd * 255
player.Color(2) = Rnd * 255
walls(20).Vector(0, 0, 0) = nl - (cWidth / 2)
walls(20).Vector(0, 0, 1) = nv + (cWidth / 2)
walls(20).Vector(1, 0, 0) = nl + (cWidth / 2)
walls(20).Vector(1, 0, 1) = nv + (cWidth / 2)
walls(20).Vector(2, 0, 0) = nl - (cWidth / 2)
walls(20).Vector(2, 0, 1) = nv + (cWidth / 2)
walls(20).Vector(3, 0, 0) = nl + (cWidth / 2)
walls(20).Vector(3, 0, 1) = nv + (cWidth / 2)
walls(20).Vector(4, 0, 0) = nl - (cWidth / 2)
walls(20).Vector(4, 0, 1) = nv - (cWidth / 2)
walls(20).Vector(5, 0, 0) = nl + (cWidth / 2)
walls(20).Vector(5, 0, 1) = nv - (cWidth / 2)
walls(20).Vector(6, 0, 0) = nl - (cWidth / 2)
walls(20).Vector(6, 0, 1) = nv - (cWidth / 2)
walls(20).Vector(7, 0, 0) = nl + (cWidth / 2)
walls(20).Vector(7, 0, 1) = nv - (cWidth / 2)

For i = 19 To 0 Step -1
walls(i).Vector(0, 0, 0) = walls(i - 0).Vector(2, 0, 0)
walls(i).Vector(0, 0, 1) = walls(i - 0).Vector(2, 0, 1)
walls(i).Vector(1, 0, 0) = walls(i - 0).Vector(3, 0, 0)
walls(i).Vector(1, 0, 1) = walls(i - 0).Vector(3, 0, 1)
walls(i).Vector(2, 0, 0) = walls(i + 1).Vector(0, 0, 0)
walls(i).Vector(2, 0, 1) = walls(i + 1).Vector(0, 0, 1)
walls(i).Vector(3, 0, 0) = walls(i + 1).Vector(1, 0, 0)
walls(i).Vector(3, 0, 1) = walls(i + 1).Vector(1, 0, 1)
walls(i).Vector(4, 0, 0) = walls(i - 0).Vector(6, 0, 0)
walls(i).Vector(4, 0, 1) = walls(i - 0).Vector(6, 0, 1)
walls(i).Vector(5, 0, 0) = walls(i - 0).Vector(7, 0, 0)
walls(i).Vector(5, 0, 1) = walls(i - 0).Vector(7, 0, 1)
walls(i).Vector(6, 0, 0) = walls(i + 1).Vector(4, 0, 0)
walls(i).Vector(6, 0, 1) = walls(i + 1).Vector(4, 0, 1)
walls(i).Vector(7, 0, 0) = walls(i + 1).Vector(5, 0, 0)
walls(i).Vector(7, 0, 1) = walls(i + 1).Vector(5, 0, 1)

Next i


Me.Cls
player.y = player.y + 2
If up Then
player.y = player.y - 10
 player.xTheta = 30
End If
If dw Then
player.y = player.y + 10
 player.xTheta = 330
End If
If le Then
player.x = player.x - 10
 player.zTheta = 330
End If
If ri Then
player.x = player.x + 10
player.zTheta = 30
End If
If ri = False And le = False Then
player.zTheta = 0
End If
If up = False And dw = False Then
player.xTheta = 0
End If
For i = 20 To 0 Step -1
walls(i).Render Me
Next i
player.Render Me
BitBlt prim.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hdc, 0, 0, vbSrcCopy
Me.ForeColor = vbBlack
TextOut Me.hdc, 10, 10, ("Score: " & score), 100
If cWidth > 20 Then
cWidth = cWidth - 0.5
wall = wall - 0.5
ceil = ceil - 0.5
End If
DoEvents
score = score + (10000 / cWidth)
Debug.Print walls(0).Vector(1, 0, 0) & walls(0).Vector(4, 0, 1)
If player.x < -walls(1).Vector(1, 0, 0) Or player.x > -walls(1).Vector(0, 0, 0) Or player.y < -walls(1).Vector(1, 0, 1) Or player.y > -walls(1).Vector(4, 0, 1) Then
MsgBox "You lose with score of " & score
cmdStart.visible = True
lblTitle(0).visible = True
lblTitle(1).visible = True
lblName.visible = True
prim.Cls
Call Form_Load
End If

If running = False Then
tmr1.Enabled = False
End If
End Sub
