VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sdPatch As DDSURFACEDESC2
Public dsPatch As DirectDrawSurface7
Private recPatch As RECT

Private m_XPos As Long
Private m_YPos As Long

Private Sub Class_Initialize()
Dim Key As DDCOLORKEY

sdPatch.lFlags = DDSD_CAPS
sdPatch.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set dsPatch = ddMain.CreateSurfaceFromFile(App.Path & "\Patch.BMP", sdPatch)
recPatch.Bottom = sdPatch.lHeight
recPatch.Right = sdPatch.lWidth
Key.low = 0
Key.high = 0
dsPatch.SetColorKey DDCKEY_SRCBLT, Key

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
RECTO = recPatch
End Property

Public Sub ShiftDown()
YPos = YPos + SPEED + ShiftSpeed
If YPos > 480 Then
    YPos = -100
    XPos = Int(((200 + Road.RECTO.Right - RECTO.Right) - 200 + 1) * Rnd + 200)
End If
End Sub

