VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sdCar As DDSURFACEDESC2
Public dsCar As DirectDrawSurface7
Private recCar As RECT

Private m_XPos
Private m_YPos

Private Sub Class_Initialize()
Dim Key As DDCOLORKEY

sdCar.lFlags = DDSD_CAPS
sdCar.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set dsCar = ddMain.CreateSurfaceFromFile(App.Path & "\car.BMP", sdCar)
recCar.Bottom = sdCar.lHeight
recCar.Right = sdCar.lWidth
Key.low = 0
Key.high = 0
dsCar.SetColorKey DDCKEY_SRCBLT, Key

XPos = 240
YPos = 300
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

Public Property Get RECTO() As RECT
RECTO = recCar
End Property

Public Property Get DestRect() As RECT
DestRect.Bottom = YPos + RECTO.Bottom
DestRect.Top = YPos
DestRect.Left = XPos
DestRect.Right = XPos + RECTO.Right
End Property


Public Function ShiftRight()
If Not m_XPos >= 340 Then
    m_XPos = m_XPos + 5
End If
End Function

Public Function ShiftLeft()
If Not m_XPos <= 205 Then m_XPos = m_XPos - 5
End Function

Public Function ShiftUP()
If Not m_YPos <= 0 Then
    m_YPos = m_YPos - 5
Else: m_YPos = 0
End If
End Function

Public Function ShiftDown()
If Not m_YPos >= 480 - RECTO.Bottom Then
    m_YPos = m_YPos + 5
Else
    m_YPos = 480 - RECTO.Bottom
End If
End Function


