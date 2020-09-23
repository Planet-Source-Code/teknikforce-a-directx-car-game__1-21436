Attribute VB_Name = "Variables"
Option Explicit

'The DirectDraw Declares
Public DXMain As New DirectX7 'The Main DirectX Object
Public ddMain As DirectDraw7 'The Main DirectDraw Object
Public diMain As DirectInput 'The Main DirectInput Object
Public diDev As DirectInputDevice 'The DirectX Input Device for DirectInput
Public diState As DIKEYBOARDSTATE 'The Enum That stores the keyboard state for the diDevi
Public sdMain As DDSURFACEDESC2 'This type describes the main surface
Public sdBack As DDSURFACEDESC2 'The Backbuffer Surface
Public dsPrim As DirectDrawSurface7 'The Main Surface
Public dsBbuf As DirectDrawSurface7 'The Back Buffer Surface
Public dsClip As DirectDrawClipper 'The DirectDraw Clipper Object

Public Road As CRoad
Public Car As CCar
Public Guide(3) As CGuide
Public Patch As CPatch
Public Feed As CFeed
Public Fuel As CFuel

Public FPSTotal As Long
Public FPS As Long

Public ShiftSpeed As Single

Public Const SPEED = 4

Public Score As Long
