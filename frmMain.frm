VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "frmMain"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   295
   ScaleMode       =   0  'User
   ScaleWidth      =   488.846
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1320
      Top             =   2520
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
    DxEnd
End If
End Sub

Private Sub Form_Load()
ShowCursor 0 'Hides The Cursor
InitDx 'Initializes Directx
INITVars 'Initialize some variables
Game_Loop 'Every Game `Has One of This
End Sub

Public Sub InitDx()
Dim rc(1) As RECT
Set ddMain = DXMain.DirectDrawCreate("") 'Set the ddMain object to a new instance of DirectDraw

Me.Show 'Shows the form

ddMain.SetCooperativeLevel Me.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN 'We want to monopolize the computer and display the thing in full screen
ddMain.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT 'sets the screen size and bit depth, DDSM_DDEFAULT is same in all projects

'Describe the Primary Surface
sdMain.lFlags = DDSD_CAPS 'Enables the DDSD_CAPS member
sdMain.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE  'This is a primary surface, is a complex surface and can be flipped

'Describes the back buffer
sdBack.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY
sdBack.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
sdBack.lWidth = 640
sdBack.lHeight = 480

Set dsPrim = ddMain.CreateSurface(sdMain) 'Make a new surface that fits our description
Set dsBbuf = ddMain.CreateSurface(sdBack) 'Mack the back buffer

Set dsClip = ddMain.CreateClipper(0) 'Create the clipper object
dsClip.SetHWnd frmMain.hWnd
dsBbuf.SetClipper dsClip
dsPrim.SetClipper dsClip

'The DirectInput Initialization
Set diMain = DXMain.DirectInputCreate 'Create an instance of DirectInput
Set diDev = diMain.CreateDevice("GUID_SysKeyboard") 'Create an instance of the device
diDev.SetCommonDataFormat DIFORMAT_KEYBOARD 'Set data format to keyboard
diDev.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE 'Set Co-operative level to Background and non_exclusive to give other applications some room
diDev.Acquire 'Get The Device

dsBbuf.SetForeColor vbWhite   'Set the forecolor to white to draw text
dsBbuf.SetFont Me.Font  'Set the font to the form font
End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowCursor 1
End Sub

Public Sub Game_Loop()
Do
    Do_Keys 'Check out the key action
    DxBlit 'Blit The Sprites
    DoEvents
Loop
End Sub

Public Sub Do_Keys()
diDev.GetDeviceStateKeyboard diState

If diState.Key(DIK_ESCAPE) <> 0 Then
    DxEnd
End If

'Check right key press
If diState.Key(DIK_RIGHT) <> 0 Then
    Car.ShiftRight
End If

'Check left key press
If diState.Key(DIK_LEFT) <> 0 Then
    Car.ShiftLeft
End If

'Check Up key press
If diState.Key(DIK_UP) <> 0 Then
    Car.ShiftUP
End If

'Check left key press
If diState.Key(DIK_DOWN) <> 0 Then
    Car.ShiftDown
End If
End Sub

Private Sub DxBlit()
Dim rect2 As RECT
Dim DestRect As RECT

FPS = FPS + 1
dsBbuf.SetFillColor Me.BackColor
'dsBbuf.SetForeColor Me.BackColor
dsBbuf.DrawBox 0, 0, 640, 480
dsBbuf.SetForeColor vbWhite
dsBbuf.DrawText 540, 440, "By: Cyril M Gupta", False
dsBbuf.DrawText 540, 460, "FPS: " & FPSTotal, False
dsBbuf.DrawText 10, 100, "Score: " & Score, False
dsBbuf.DrawText 10, 150, "Speed: " & SPEED + ShiftSpeed, False
dsBbuf.DrawText 540, 460, "FPS: " & FPSTotal, False

dsBbuf.SetForeColor Me.ForeColor

'Blit The Road
'With Road
'    dsBbuf.Blt .DestRect, .dsRoad, .RECTO, DDBLT_WAIT Or DDBLT_KEYSRC
'End With

''Blit The Guides
With Guide(0)
    dsBbuf.Blt .DestRect, .dsGuide, .RECTO, DDBLT_KEYSRC Or DDBLT_WAIT
    .ShiftDown
End With
With Guide(1)
    dsBbuf.Blt .DestRect, .dsGuide, .RECTO, DDBLT_KEYSRC Or DDBLT_WAIT
    .ShiftDown
End With
With Guide(2)
    dsBbuf.Blt .DestRect, .dsGuide, .RECTO, DDBLT_KEYSRC Or DDBLT_WAIT
    .ShiftDown
End With

'Fuel Logo
With Feed
    dsBbuf.Blt .DestRect, .dsFeed, .RECTO, DDBLT_KEYSRC Or DDBLT_WAIT
End With

'Blit The Patch
With Patch
    dsBbuf.Blt .DestRect, .dsPatch, .RECTO, DDBLT_KEYSRC Or DDBLT_WAIT
    .ShiftDown
End With

'Blit The Fuel
With Fuel
    dsBbuf.Blt .DestRect, .dsFuel, .RECTO, DDBLT_KEYSRC Or DDBLT_WAIT
    .ShiftDown
End With

'Blit The Car
With Car
    dsBbuf.Blt .DestRect, .dsCar, .RECTO, DDBLT_KEYSRC Or DDBLT_WAIT
End With

dsPrim.Blt DestRect, dsBbuf, rect2, DDBLT_WAIT

If Fuel.DestRect.Left >= Car.DestRect.Left And _
   Fuel.DestRect.Right <= Car.DestRect.Right And _
   Fuel.DestRect.Top >= Car.DestRect.Top And _
   Fuel.DestRect.Bottom <= Car.DestRect.Bottom Then
   Fuel.YPos = 480
   Score = Score + 300
Else
'    Debug.Print "Car: "; Car.DestRect.Left
'    Debug.Print "Feed: "; Fuel.DestRect.Left
End If

If Patch.DestRect.Left >= Car.DestRect.Left And Patch.DestRect.Right <= Car.DestRect.Right Or _
   Patch.DestRect.Left <= Car.DestRect.Right And Patch.DestRect.Right <= Car.DestRect.Right Then
    Score = Score - 10
End If

If Not (SPEED + ShiftSpeed) > 9 Then ShiftSpeed = ShiftSpeed + 0.001

End Sub

Private Sub INITVars()
Set Road = New CRoad
Set Car = New CCar
Set Guide(0) = New CGuide
Set Guide(1) = New CGuide
Set Guide(2) = New CGuide
Set Feed = New CFeed
Set Patch = New CPatch
Set Fuel = New CFuel

Guide(0).XPos = 290
Guide(0).YPos = 0

Guide(1).XPos = 290
Guide(1).YPos = 200

Guide(2).XPos = 290
Guide(2).YPos = 400
End Sub

Private Sub DxEnd()
'This sub unloads DirectX and puts control back to the computer
ShowCursor 1
'This sub unloads DirectX and puts control back to the computer
'ShowCursor 1
ddMain.RestoreDisplayMode 'Restores the old resolution
'ddMain.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL 'Setco-op to normal
diDev.Unacquire 'Disable directinput

'DoEvents
'MsgBox "That's the end." & vbCrLf & "If you liked this game vote for it."
End
End Sub

Private Sub Timer1_Timer()
FPSTotal = FPS
FPS = 0
'Debug.Print FPSTotal
End Sub
