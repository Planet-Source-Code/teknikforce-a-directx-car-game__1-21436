VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Option Explicit

Private sdFeed As DDSURFACEDESC2
Public dsFeed As DirectDrawSurface7
Private recFeed As RECT

Private m_XPos As Long
Private m_YPos As Long

Private Sub Class_Initialize()
Dim Key As DDCOLORKEY

sdFeed.lFlags = DDSD_CAPS
sdFeed.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set dsFeed = ddMain.CreateSurfaceFromFile(App.Path & "\Need.BMP", sdFeed)
recFeed.Bottom = sdFeed.lHeight
recFeed.Right = sdFeed.lWidth
Key.low = 0
Key.high = 0
dsFeed.SetColorKey DDCKEY_SRCBLT, Key

XPos = 10
YPos = 10
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
RECTO = recFeed
End Property



