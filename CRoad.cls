VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sdRoad As DDSURFACEDESC2 'The description of the road surface
Public dsRoad As DirectDrawSurface7 'The Surface of The Road
Private RectRoad As RECT

Public Function XPos()
XPos = 200
End Function

Public Function YPos()
YPos = 0
End Function

Public Property Get RECTO() As RECT
RECTO = RectRoad
End Property

Public Property Get DestRect() As RECT
DestRect.Bottom = 480
DestRect.Top = 0
DestRect.Left = 200
DestRect.Right = 200 + RECTO.Right
End Property

Private Sub Class_Initialize()
Dim Key As DDCOLORKEY

sdRoad.lFlags = DDSD_CAPS
sdRoad.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set dsRoad = ddMain.CreateSurfaceFromFile(App.Path & "\ROAD.BMP", sdRoad)
RectRoad.Bottom = sdRoad.lHeight
RectRoad.Right = sdRoad.lWidth
Key.low = 0
Key.high = 0
dsRoad.SetColorKey DDCKEY_SRCBLT, Key
End Sub
