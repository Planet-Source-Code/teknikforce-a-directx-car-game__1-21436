VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFuel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Option Explicit

Private sdFuel As DDSURFACEDESC2
Public dsFuel As DirectDrawSurface7
Private recFuel As RECT

Private m_XPos As Long
Private m_YPos As Long

Private Sub Class_Initialize()
Dim Key As DDCOLORKEY

sdFuel.lFlags = DDSD_CAPS
sdFuel.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set dsFuel = ddMain.CreateSurfaceFromFile(App.Path & "\Feed.BMP", sdFuel)
recFuel.Bottom = sdFuel.lHeight
recFuel.Right = sdFuel.lWidth
Key.low = 0
Key.high = 0
dsFuel.SetColorKey DDCKEY_SRCBLT, Key

XPos = 290
YPos = 100
End Sub

Public Property Get XPos() As Long
XPos = m_XPos
End Property
Public Property Let XPos(vNewVal As Long)
m_XPos = vNewVal
End Property

Public Property Get YPos() As Long
YPos = m_YPos
End Property
Public Property Let YPos(vNewVal As Long)
m_YPos = vNewVal
End Property

Public Property Get DestRect() As RECT
DestRect.Bottom = YPos + RECTO.Bottom
DestRect.Right = XPos + RECTO.Right
DestRect.Top = YPos
DestRect.Left = XPos
End Property

Public Property Get RECTO() As RECT
RECTO = recFuel
End Property

Public Sub ShiftDown()
YPos = YPos + SPEED + ShiftSpeed
If YPos > 480 Then
    YPos = -100
    Randomize
    XPos = Int(((200 + Road.RECTO.Right - RECTO.Right) - 200 + 1) * Rnd + 200)
End If
End Sub



