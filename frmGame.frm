VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   295
   ScaleMode       =   0  'User
   ScaleWidth      =   345.191
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Main DirectX object
Dim Dx As New DirectX7
'Maub DirectDraw object
Dim DDraw As DirectDraw7

'Surfaces
'Primary surface
Dim ddsPrimary As DirectDrawSurface7
'Backbuffer
Dim ddsBackBuffer As DirectDrawSurface7

'Background
Dim Background(2, 2) As DirectDrawSurface7
Dim s1024x768rect(2, 2) As RECT 'RECT variable
'Sprite
Dim Sprite As DirectDrawSurface7
Dim SpriteRECT As RECT 'Sprite's RECT variable
Dim SpriteAnimRECT(8, 4) As RECT
'map edit square
Dim Square As DirectDrawSurface7
Dim SquareRECT As RECT
'Base tile set
Dim Grass As DirectDrawSurface7
Dim GrassRECT As RECT
'Object tile set
Dim Tree As DirectDrawSurface7
Dim TreeRECT As RECT


'Surface Descriptions
Dim ddsdPrimary As DDSURFACEDESC2
Dim ddsdBackbuffer As DDSURFACEDESC2

'font
Dim fFont As New StdFont

'DirectInput interface
Dim DInput As DirectInput

'Input devices
Dim DI_Mouse As DirectInputDevice
Dim DI_KeyBoard As DirectInputDevice

'These variables will be filled with device info
Dim KeyboardState As DIKEYBOARDSTATE
Dim MouseState As DIMOUSESTATE

'character and screen position
Dim charX As Integer
Dim charY As Integer
Dim scrX As Integer
Dim scrY As Integer
Dim charXtmp As Integer
Dim charYtmp As Integer
Dim scrXtmp As Integer
Dim scrYtmp As Integer

'animation varible
Dim directanim As Integer
Dim countanim As Single

'key press varibles
Dim kUP As Boolean
Dim kDOWN As Boolean
Dim kLEFT As Boolean
Dim kRIGHT As Boolean
Dim didmove As Boolean

'edit varibles
Dim gamemode As Integer
Dim presetobj() As Integer
Dim presetnames() As String

'map varibles
Dim map(95, 71, 2, 1) As Integer
Dim d1 As Integer
Dim d2 As Integer

'timing varibles
Dim LastTimeChecked As Long
Const DelayTime As Integer = 1000 / 30
Const DelayTimeEM As Integer = 1000 / 10

Dim running As Boolean

Private Sub Form_Click()
    'If user clicks on form end the program
    running = False
End Sub

Private Sub Form_Load()
charX = 480
charY = 352
scrX = 0
scrY = 0

init
gameloop
End Sub

Private Sub init()
    running = True 'the app is running
      
    'Initializing DirectDraw
    'On Error Resume Next
    

    'Create Directdraw
    Set DDraw = Dx.DirectDrawCreate("")
      
    frmGame.Show 'show the form
    
    'Setting cooperative level and display mode
    'Fullscreen exclusive mode
    DDraw.SetCooperativeLevel frmGame.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
    '640x480 resolution in 16-bit colour (about 16 000 colours)
    'This is the most common BPP (Bits-Per-Pixel) used today.
    DDraw.SetDisplayMode 1024, 768, 24, 0, DDSDM_DEFAULT


    
    'Fill out Primary Surface Description
    ddsdPrimary.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    'Using Video Memory for the primary surface and backbuffer
    ddsdPrimary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
      
    ddsdPrimary.lBackBufferCount = 1 'One Backbuffer
      
    'Create the Primary Surface
    Set ddsPrimary = DDraw.CreateSurface(ddsdPrimary)


    'Backbuffer
    Dim Caps As DDSCAPS2
    Caps.lCaps = DDSCAPS_BACKBUFFER
    
    
    Set ddsBackBuffer = ddsPrimary.GetAttachedSurface(Caps)
    'Fill out description
    ddsBackBuffer.GetSurfaceDesc ddsdBackbuffer

    'font
    fFont.Name = "Arial"
    fFont.Size = 16
    fFont.Bold = True
    ddsBackBuffer.SetFont fFont
    ddsBackBuffer.SetForeColor vbBlack
    ddsBackBuffer.SetFontTransparency True
    
    'Load bitmap files into surfaces
    For aa = 0 To 2
    For bb = 0 To 2
    DDCreateSurface Background(aa, bb), App.Path & "\background.bmp", s1024x768rect(aa, bb)
    Next bb
    Next aa
    DDCreateSurface Sprite, App.Path & "\char.bmp", SpriteRECT
    getanimrects SpriteAnimRECT, SpriteRECT, 8, 4
    DDCreateSurface Square, App.Path & "\sq.bmp", SquareRECT
    DDCreateSurface Grass, App.Path & "\tiles.bmp", GrassRECT
    DDCreateSurface Tree, App.Path & "\objs.bmp", TreeRECT
    mapfill App.Path & "\map2.map"
    
'Create DirectInput interface from the DirectX7 interface
Set DInput = Dx.DirectInputCreate

'Create mouse and keyboard device, set dataformat, set device cooperative level and acquire device
Set DI_Mouse = DInput.CreateDevice("GUID_SysMouse")
DI_Mouse.SetCommonDataFormat DIFORMAT_MOUSE
DI_Mouse.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
DI_Mouse.Acquire

Set DI_KeyBoard = DInput.CreateDevice("GUID_SysKeyboard")
DI_KeyBoard.SetCommonDataFormat DIFORMAT_KEYBOARD
DI_KeyBoard.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
DI_KeyBoard.Acquire
End Sub

Private Sub gameloop()
Dim jumpsize As Integer
Dim DTtmp As Integer
gamemode = 0

    'Render loop
    Do
        If gamemode = 0 Then DTtmp = DelayTime: jumpsize = 8
        If gamemode = 1 Then DTtmp = DelayTimeEM: jumpsize = 32
        If Dx.TickCount - LastTimeChecked >= DTtmp Then
            LastTimeChecked = Dx.TickCount
            DoEvents
            jumpsize = 8
            didmove = False
            'Colorfill the background with flat black
            kUP = True
            kDOWN = True
            kRIGHT = True
            kLEFT = True
            If gamemode = 0 Then contactdetect
            DI_KeyBoard.GetDeviceStateKeyboard KeyboardState
        
            If KeyboardState.Key(DIK_DOWN) <> 0 Then
                directanim = 0
                didmove = True
                If kDOWN = True Then
                    scrY = scrY + jumpsize
                End If
            End If
        
            If KeyboardState.Key(DIK_UP) <> 0 Then
                didmove = True
                directanim = 1
                If kUP = True Then
                    scrY = scrY - jumpsize
                End If
            End If
        
            If KeyboardState.Key(DIK_RIGHT) <> 0 Then
                didmove = True
                directanim = 2
                If kRIGHT = True Then
                    scrX = scrX + jumpsize
                End If
            End If
        
            If KeyboardState.Key(DIK_LEFT) <> 0 Then
                directanim = 3
                didmove = True
                If kLEFT = True Then
                    scrX = scrX - jumpsize
                End If
            End If
        
            If KeyboardState.Key(DIK_ESCAPE) <> 0 Then running = False
            If KeyboardState.Key(DIK_F2) <> 0 Then gamemode = 1: editinit
        
            If didmove = False Then countanim = 0
            If didmove = True Then countanim = countanim + 1
            If countanim = 9 Then countanim = 1
        
            ddsBackBuffer.BltColorFill s1024x768rect(0, 0), RGB(31, 0, 31)
            gameblt
        
            If gamemode = 1 Then Call ddsBackBuffer.DrawText(0, 0, Int((scrX + charX + 16) / 32) & ", " & Int((scrY + charY + 32) / 32), False)
            'Use BltFast to draw the Sprite (transparent)
            If gamemode = 0 Then Call ddsBackBuffer.BltFast(charXtmp, charYtmp, Sprite, SpriteAnimRECT(Int(countanim), directanim), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            If gamemode = 1 Then
                charXtmp = (Int((charXtmp + 32) / 32)) * 32
                charYtmp = (Int((charYtmp + 16) / 32)) * 32
                Call ddsBackBuffer.BltFast(charXtmp, charYtmp, Square, SquareRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        'Flip the backbuffer to the primary surface
            ddsPrimary.Flip Nothing, DDFLIP_WAIT
        End If
    Loop Until running = False


    'Unloading
    'With DirectDraw you must unload all your surfaces etc.
    'to clear up memory.
    For aa = 0 To 2
    For bb = 0 To 2
    Set Background(aa, bb) = Nothing
    Next bb
    Next aa
    Set Sprite = Nothing
      
    Set ddsPrimary = Nothing
    Set ddsBackBuffer = Nothing
      
    DDraw.RestoreDisplayMode
    'Restore to normal display mode
    DDraw.SetCooperativeLevel frmGame.hWnd, DDSCL_NORMAL
    End
End Sub

Private Sub DDCreateSurface(surface As DirectDrawSurface7, BmpPath As String, RECTvar As RECT, Optional TransCol As Long = 16711935, Optional UseSystemMemory As Boolean = True)
    'This sub will load a bitmap from a file
    'into a specified dd surface. Transparent
    'colour is black (0) by default.
      
    Dim tempddsd As DDSURFACEDESC2
      
    Set surface = Nothing
      
    'Load sprite
    tempddsd.lFlags = DDSD_CAPS
    If UseSystemMemory = True Then
        tempddsd.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY Or DDSCAPS_OFFSCREENPLAIN
    Else
        tempddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    End If
    Set surface = DDraw.CreateSurfaceFromFile(BmpPath, tempddsd)
      
    'set the RECT dimensions
    RECTvar.Right = tempddsd.lWidth
    RECTvar.Bottom = tempddsd.lHeight
      
    'Colour key
    Dim ddckColourKey As DDCOLORKEY
    ddckColourKey.low = TransCol
    ddckColourKey.high = TransCol
    surface.SetColorKey DDCKEY_SRCBLT, ddckColourKey
      
      
End Sub

Public Sub DDBltFast(surface As DirectDrawSurface7, destsurface As DirectDrawSurface7, RECTvar As RECT, x As Integer, y As Integer)
'mental note, remove this sub
'CLIPPING
     
    'Set up screen rect for clipping
    Dim ScreenRECT As RECT
    ScreenRECT.Top = y
    ScreenRECT.Left = x
    ScreenRECT.Bottom = y + RECTvar.Bottom
    ScreenRECT.Right = x + RECTvar.Right
     
    'Temporary rect
    Dim RectTEMP As RECT
    RectTEMP = RECTvar
     
    'Clip surface
    With ScreenRECT
        If .Bottom > 768 Then
            RectTEMP.Bottom = RectTEMP.Bottom - (.Bottom - 768)
            .Bottom = 768
        End If
        If .Left < 0 Then
            RectTEMP.Left = RectTEMP.Left - .Left
            .Left = 0
        End If
        If .Right > 1024 Then
            RectTEMP.Right = RectTEMP.Right - (.Right - 1024)
            .Right = 1024
        End If
        If .Top < 0 Then
            RectTEMP.Top = RectTEMP.Top - .Top
            .Top = 0
        End If
    End With

    Call destsurface.BltFast(ScreenRECT.Left, ScreenRECT.Top, surface, RectTEMP, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
End Sub

Private Sub mapfill(mappath As String)
Dim fL

Open mappath For Input As #1

    'first line read
    Line Input #1, fL
    Do
        If Left(fL, 1) = "*" Then
            fL = Right(fL, Len(fL) - 1)
            GoTo Wdth
        Else
            d1 = d1 & Left(fL, 1)
            fL = Right(fL, Len(fL) - 1)
        End If
    Loop
Wdth:
    Do
        If Left(fL, 1) = "*" Then
            fL = Right(fL, Len(fL) - 1)
            GoTo Hght
        Else
            d2 = d2 & Left(fL, 1)
            fL = Right(fL, Len(fL) - 1)
        End If
    Loop
Hght:
'reading the guts of the map
'ReDim map(d1 - 1, d2 - 1, 2, 1) As Integer

Dim temphold As String
Dim xx As Integer
Dim yy As Integer
Dim zz As Integer
Line Input #1, temphold
For yy = 0 To d2 - 1
    For xx = 0 To d1 - 1
        For zz = 0 To 2
            map(xx, yy, zz, 0) = Asc(Left(temphold, 1)) - 48
            temphold = Right(temphold, Len(temphold) - 1)
            map(xx, yy, zz, 1) = Asc(Left(temphold, 1)) - 48
            temphold = Right(temphold, Len(temphold) - 1)
            If map(xx, yy, zz, 0) = 0 And map(xx, yy, zz, 1) = 2 Then starthere xx, yy
        Next zz
    Next xx
Next yy
Close #1

'base & object layer print
Dim r1 As RECT
Dim r2 As RECT
For aa = 0 To 2
    For bb = 0 To 2
        For cc = 0 To 23
            ff = cc + (aa * 24)
            For dd = 0 To 31
                ee = dd + (bb * 32)
                r1.Left = map(ee, ff, 0, 0) * 32
                r1.Top = map(ee, ff, 0, 1) * 32
                r2.Left = map(ee, ff, 1, 0) * 32
                r2.Top = map(ee, ff, 1, 1) * 32
                r1.Right = r1.Left + 32
                r1.Bottom = r1.Top + 32
                r2.Right = r2.Left + 32
                r2.Bottom = r2.Top + 32
                Call Background(bb, aa).BltFast(dd * 32, cc * 32, Grass, r1, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Call Background(bb, aa).BltFast(dd * 32, cc * 32, Tree, r2, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Next dd
        Next cc
    Next bb
Next aa
End Sub

Private Sub gameblt()
Dim Xincl(2) As Boolean
Dim Yincl(2) As Boolean
Dim rect4(3) As RECT
Dim rectdest As RECT
Dim holder(3, 1) As Integer
charXtmp = charX
charYtmp = charY
scrXtmp = scrX
scrYtmp = scrY

'if the screen co-ords reach the edge of the map
'move the character instead of the screen
If scrX < 0 Then
    charXtmp = charX + scrX
    scrXtmp = 0
End If
If scrX + 1024 > d1 * 32 Then
    charXtmp = charX + ((scrX + 1024) - (d1 * 32))
    scrXtmp = (d1 * 32) - 1024
End If
If scrY < 0 Then
    charYtmp = charY + scrY
    scrYtmp = 0
End If
If scrY + 768 > d2 * 32 Then
    charYtmp = charY + ((scrY + 768) - (d2 * 32))
    scrYtmp = (d2 * 32) - 768
End If

'makes sure the charXtmp and charYtmp varibles are between 0 and 1024/768
Do
If charXtmp < 1024 Then Exit Do
charXtmp = charXtmp - 1024
Loop
Do
If charYtmp < 768 Then Exit Do
charYtmp = charYtmp - 768
Loop

'this makes sure the character stops when it reaches the edge of the screen
If gamemode = 0 Then
    If charXtmp < 0 Then charXtmp = 0: scrX = charXtmp - charX
    If charXtmp > 960 Then charXtmp = 960: scrX = (charXtmp - charX) + (d1 * 32) - 1024
    If charYtmp < 0 Then charYtmp = 0: scrY = charYtmp - charY
    If charYtmp > 704 Then charYtmp = 704: scrY = (charYtmp - charY) + (d2 * 32) - 768
ElseIf gamemode = 1 Then
    If charXtmp < -16 Then charXtmp = -16: scrX = charXtmp - charX
    If charXtmp > 944 Then charXtmp = 944: scrX = (charXtmp - charX) + (d1 * 32) - 1024
    If charYtmp < 0 Then charYtmp = 0: scrY = charYtmp - charY
    If charYtmp > 704 Then charYtmp = 704: scrY = (charYtmp - charY) + (d2 * 32) - 768
End If

'this checks which of the 9 map surfaces will be blted
If scrXtmp < 1024 Then Xincl(0) = True
If scrXtmp <> 0 And scrXtmp <> 2048 Then Xincl(1) = True
If scrXtmp > 1024 Then Xincl(2) = True
If scrYtmp < 768 Then Yincl(0) = True
If scrYtmp <> 0 And scrYtmp <> 1536 Then Yincl(1) = True
If scrYtmp > 768 Then Yincl(2) = True

'makes sure the scrXtmp and scrYtmp varibles are between 0 and 1024/768
Do
If scrXtmp < 1024 Then Exit Do
scrXtmp = scrXtmp - 1024
Loop
Do
If scrYtmp < 768 Then Exit Do
scrYtmp = scrYtmp - 768
Loop
If gamemode = 1 Then
    scrXtmp = Int(scrXtmp / 32) * 32
    scrYtmp = Int(scrYtmp / 32) * 32
End If
'calculates the 4 rectangles needed to make up the current map screen
cc = 0
For aa = 0 To 2
    For bb = 0 To 2
        If Xincl(bb) = True And Yincl(aa) = True And cc = 0 Then
            rect4(cc).Top = scrYtmp
            rect4(cc).Left = scrXtmp
            rect4(cc).Bottom = 768
            rect4(cc).Right = 1024
            holder(cc, 0) = bb
            holder(cc, 1) = aa
            cc = cc + 1
            If scrX <= 0 Then cc = 2
        ElseIf Xincl(bb) = True And Yincl(aa) = True And cc = 1 Then
            rect4(cc).Top = scrYtmp
            rect4(cc).Left = 0
            rect4(cc).Bottom = 768
            rect4(cc).Right = scrXtmp
            holder(cc, 0) = bb
            holder(cc, 1) = aa
            cc = cc + 1
        ElseIf Xincl(bb) = True And Yincl(aa) = True And cc = 2 Then
            rect4(cc).Top = 0
            rect4(cc).Left = scrXtmp
            rect4(cc).Bottom = scrYtmp
            rect4(cc).Right = 1024
            holder(cc, 0) = bb
            holder(cc, 1) = aa
            cc = cc + 1
        ElseIf Xincl(bb) = True And Yincl(aa) = True And cc = 3 Then
            rect4(cc).Top = 0
            rect4(cc).Left = 0
            rect4(cc).Bottom = scrYtmp
            rect4(cc).Right = scrXtmp
            holder(cc, 0) = bb
            holder(cc, 1) = aa
            Exit For
        End If
    Next bb
Next aa

'
rectdest.Left = 1024 - (rect4(1).Right - rect4(1).Left)
rectdest.Top = 768 - (rect4(2).Bottom - rect4(2).Top)

'blting finally
ddsBackBuffer.BltFast 0, 0, Background(holder(0, 0), holder(0, 1)), rect4(0), DDBLTFAST_WAIT
ddsBackBuffer.BltFast rectdest.Left, 0, Background(holder(1, 0), holder(1, 1)), rect4(1), DDBLTFAST_WAIT
ddsBackBuffer.BltFast 0, rectdest.Top, Background(holder(2, 0), holder(2, 1)), rect4(2), DDBLTFAST_WAIT
ddsBackBuffer.BltFast rectdest.Left, rectdest.Top, Background(holder(3, 0), holder(3, 1)), rect4(3), DDBLTFAST_WAIT

End Sub

Public Sub getanimrects(rect1() As RECT, rect2 As RECT, itemsX As Integer, itemsY As Integer)
Dim div1 As Integer
Dim div2 As Integer
div1 = (rect2.Right - rect2.Left) / (itemsX + 1)
div2 = (rect2.Bottom - rect2.Top) / (itemsY + 1)
For aa = 0 To itemsY
    For bb = 0 To itemsX
        rect1(bb, aa).Left = bb * div1
        rect1(bb, aa).Right = rect1(bb, aa).Left + div1
        rect1(bb, aa).Top = aa * div2
        rect1(bb, aa).Bottom = rect1(bb, aa).Top + div2
    Next bb
Next aa
End Sub

Public Sub starthere(screenX As Integer, screenY As Integer)
Dim tempX As Integer
Dim tempY As Integer
tempX = screenX * 32
tempY = screenY * 32
tempY = tempY + 16
scrX = tempX - 512
scrY = tempY - 384
End Sub

Public Sub contactdetect()
Dim trueposX As Integer
Dim trueposY As Integer
trueposX = scrX + charX + 16
trueposY = scrY + charY + 32

jj = 1
'down & up
If trueposY = Int(trueposY / 32) * 32 Then
    If trueposX = Int(trueposX / 32) * 32 Then jj = 0
    For ii = 0 To jj
        If map(Int(trueposX / 32) + ii, (trueposY / 32) + 1, 2, 0) = 0 Then
            If map(Int(trueposX / 32) + ii, (trueposY / 32) + 1, 2, 1) = 1 Then
                kDOWN = False
            End If
        End If
        If map(Int(trueposX / 32) + ii, (trueposY / 32) - 1, 2, 0) = 0 Then
            If map(Int(trueposX / 32) + ii, (trueposY / 32) - 1, 2, 1) = 1 Then
                kUP = False
            End If
        End If
    Next ii
End If

jj = 1
'right & left
If trueposX = Int(trueposX / 32) * 32 Then
    If trueposY = Int(trueposY / 32) * 32 Then jj = 0
    For ii = 0 To jj
        If map((trueposX / 32) + 1, Int(trueposY / 32) + ii, 2, 0) = 0 Then
            If map((trueposX / 32) + 1, Int(trueposY / 32) + ii, 2, 1) = 1 Then
                kRIGHT = False
            End If
        End If
        If map((trueposX / 32) - 1, Int(trueposY / 32) + ii, 2, 0) = 0 Then
            If map((trueposX / 32) - 1, Int(trueposY / 32) + ii, 2, 1) = 1 Then
                kLEFT = False
            End If
        End If
    Next ii
End If

End Sub

Public Sub editinit()
Dim tempstr As String
Dim tempnum(1) As Integer
Dim tempset As String
Dim n As String
Open App.Path & "\object.set" For Input As #1
Line Input #1, n
ReDim presetobj(Val(n) - 1, 9, 9, 3)
ReDim presetnames(Val(n) - 1)
For hh = 0 To Val(n) - 1
    Line Input #1, presetnames(hh)
    Line Input #1, tempstr
    Line Input #1, tempset
    tempnum(0) = Asc(Left(tempstr, 1)) - 48
    tempnum(1) = Asc(Right(tempstr, 1)) - 48
    For yy = 0 To tempnum(1) - 1
        For xx = 0 To tempnum(0) - 1
            presetobj(hh, xx, yy, 0) = Asc(Left(tempset, 1)) - 48
            tempset = Right(tempset, Len(tempset) - 1)
            presetobj(hh, xx, yy, 1) = Asc(Left(tempset, 1)) - 48
            tempset = Right(tempset, Len(tempset) - 1)
            presetobj(hh, xx, yy, 2) = Asc(Left(tempset, 1)) - 48
            tempset = Right(tempset, Len(tempset) - 1)
            presetobj(hh, xx, yy, 3) = Asc(Left(tempset, 1)) - 48
            tempset = Right(tempset, Len(tempset) - 1)
        Next xx
    Next yy
Next hh
Close #1
End Sub
