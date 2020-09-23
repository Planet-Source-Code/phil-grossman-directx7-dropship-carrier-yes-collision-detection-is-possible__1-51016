VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   600
      Top             =   2040
   End
   Begin VB.Timer Timer6 
      Interval        =   100
      Left            =   3360
      Top             =   120
   End
   Begin VB.Timer Timer5 
      Left            =   2580
      Top             =   120
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   1980
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1395
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   810
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   195
      Top             =   135
   End
   Begin VB.Label Label1 
      Caption         =   "Loading..............."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   10935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'General Declarations
Dim pos As D3DVECTOR
Dim bulletpos As D3DVECTOR
Dim Bulletfire As D3DVECTOR
Dim RAMPXYZ1 As D3DVECTOR
Dim RAMPXYZ2 As D3DVECTOR
Dim AIRvec1 As D3DVECTOR
Dim AIRvec2 As D3DVECTOR
Dim ENvec As D3DVECTOR
Dim SOLvec As D3DVECTOR
'These three components are the big daddy of what your 3D Program.
Dim DX_Main As New DirectX7 ' The DirectX core file (The heart of it all)
Dim DD_Main As DirectDraw4 ' The DirectDraw Object
Dim D3D_Main As Direct3DRM3 ' The direct3D Object

'DirectInput Components
Dim DI_Main As DirectInput ' The DirectInput core
Dim DI_Device As DirectInputDevice ' The DirectInput device
Dim DI_State As DIKEYBOARDSTATE 'Array Holding the state of the keys

'DirectDraw Surfaces- Where the screen is drawn.
Dim DS_Front As DirectDrawSurface4 ' The frontbuffer (What you see on the screen)
Dim DS_Back As DirectDrawSurface4 ' The backbuffer, (Where everything is drawn before it's put on the screen.)
Dim SD_Front As DDSURFACEDESC2 ' The SurfaceDescription
Dim DD_Back As DDSCAPS2 ' General Surface Info

'ViewPort and Direct3D Device
Dim D3D_Device As Direct3DRMDevice3 'The Main Direct3D Retained Mode Device
Dim D3D_ViewPort As Direct3DRMViewport2 'The Direct3D Retained Mode Viewport (Kinda the camera)

'The Frames
Dim FR_Root As Direct3DRMFrame3 'The Main Frame (The other frames are put under this one (Like a tree))
Dim FR_Camera As Direct3DRMFrame3 'Another frame, just happens to be called 'camera'. We will use this
                                                   'as the viewport. Hence, the name camera. Doesn't have to be
                                                   'called 'camera'. It could be called 'BillyBob' for all I care.
Dim FR_Light(4) As Direct3DRMFrame3 'This frame contains our, guess what, spotlight!
Dim FR_Building(50) As Direct3DRMFrame3 'Frame containing our 1st mesh that will be put in this "game".
Dim FR_BULLET As Direct3DRMFrame3
Dim FR_GUN As Direct3DRMFrame3
Dim FR_BACKOB(1) As Direct3DRMFrame3
Dim FR_OBJECTS(100) As Direct3DRMFrame3

'Meshes (3D objects loaded from a .x file)
Dim MS_Building(1) As Direct3DRMMeshBuilder3
Dim MS_BULLET As Direct3DRMMeshBuilder3
Dim MS_GUN As Direct3DRMMeshBuilder3
Dim MS_BACKOB(1) As Direct3DRMMeshBuilder3
Dim MS_OBJECTS(1) As Direct3DRMMeshBuilder3

'Lights
Dim LT_Ambient As Direct3DRMLight 'Our Main (Ambient) light that illuminates everything (not just part of something
                                                'like a spotlight.
Dim LT_Spot(4) As Direct3DRMLight 'Our Spot light, makes it look more realistic.

'Camera Positions
Dim xx As Long
Dim yy As Long
Dim zz As Long

Dim esc As Boolean 'If Escape is pressed, the DX_Input sub will make it true and the main loop will end.
'Incase you haven't caught on, the prefix FR = frame, MS = Mesh, & LT = light.

Dim BackGround As Direct3DRMTexture3 'This will be the texture that holds our background.
'Note: Texture files must have side lengths that are devisible by 2!

' The following are all misc variables
Dim CameraPos As D3DVECTOR
Dim Tempread, Genc, Genc2, Genc3
Dim MODELINDEX(), MODELNAME(), MODELx(), MODELy(), MODELz(), MODELrotx(), MODELroty(), MODELrotz(), MODELTEX()
Dim Tempx As Single, Tempy As Single, Tempz As Single, KeyDown As DIKEYBOARDSTATE, COLflg
Dim LightX, LightY, LightZ, LightPen, WeaponX, Weaponz, CAMROT As Integer, OLDCAMROT As Integer
Dim LiftReached As String, CollisionON, LEVEl As Integer, LiftY, liftUP, FpS, oldfps, InfoOn
Dim Totsecs, AvFPS, totFrames, RotX, YLiftMov, Shotfired, ShotVel, MeRotx, MeRotZ, DegRot, bullX, bullZ
Dim GunMove, GunCount, CollHIT, CAMY, AMBfix, Foot, RAMPY, BAYdoor, AIRlock1, AIRlock2, AIR1, AIR2
Dim animFrame, HIT
'=============
Const cSoundBuffers = 10
'======================================================================================

Private Sub DX_Init()
On Error GoTo DXINITHELL:
 'Type, not copy and paste!
 'This sub will initialize all your components and set them up.
ShowCursor (0)  ' hide the cursor
 Set DD_Main = DX_Main.DirectDraw4Create("") 'Create the DirectDraw Object

 DD_Main.SetCooperativeLevel Form1.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE 'Set Screen Mode (Full 'Screen)
 DD_Main.SetDisplayMode 640, 480, 32, 0, DDSDM_DEFAULT 'Set Resolution and BitDepth (Lets use 32-bit color)
 
 SD_Front.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
 SD_Front.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_3DDEVICE Or DDSCAPS_COMPLEX Or _
 DDSCAPS_FLIP 'I used the line-continuation ( _ ) because the whole thing wouldn't fit on one line...
 SD_Front.lBackBufferCount = 1 'Make one backbuffer
 Set DS_Front = DD_Main.CreateSurface(SD_Front) 'Initialize the front buffer (the screen)
 'The Previous block of code just created the screen and the backbuffer.
 
 DD_Back.lCaps = DDSCAPS_BACKBUFFER
 
 Set DS_Back = DS_Front.GetAttachedSurface(DD_Back)
 DS_Back.SetForeColor RGB(255, 255, 255)
 'The backbuffer was initialized and the DirectDraw text color was set to white.

 Set D3D_Main = DX_Main.Direct3DRMCreate() 'Creates the Direct3D Retained Mode Object!!!!!

 Set D3D_Device = D3D_Main.CreateDeviceFromSurface("IID_IDirect3DHALDevice", DD_Main, DS_Back, _
 D3DRMDEVICE_DEFAULT) 'Tell the Direct3D Device that we are using hardware rendering (HALDevice) instead
                                   'of software enumeration (RGBDevice).
 D3D_Device.SetBufferCount 2 'Set the number of buffers
 D3D_Device.SetQuality D3DRMRENDER_GOURAUD 'Set Rendering Quality. Can use Flat, or WireFrame, but
                                                                  'GOURAUD has the best rendering quality.
 D3D_Device.SetTextureQuality D3DRMTEXTURE_NEAREST  'Set the texture quality
 D3D_Device.SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY 'Set the render mode.

 Set DI_Main = DX_Main.DirectInputCreate() 'Create the DirectInput Device
 Set DI_Device = DI_Main.CreateDevice("GUID_SysKeyboard") 'Set it to use the keyboard.
 DI_Device.SetCommonDataFormat DIFORMAT_KEYBOARD 'Set the data format to the keyboard format.
 DI_Device.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE 'Set Coperative Level.
 DI_Device.Acquire
 'The above block of code configures the DirectInput Device and starts it.

Exit Sub
DXINITHELL:
MsgBox Err.Description
End Sub

'=====================================================================================

Private Sub DX_MakeObjects()
On Error GoTo DXHELL:
 Set FR_Root = D3D_Main.CreateFrame(Nothing) 'This will be the root frame of the 'tree'
 With FR_Root
    .SetSceneFogMethod D3DRMFOGMETHOD_VERTEX
    .SetSceneFogMode D3DRMFOG_EXPONENTIALSQUARED
    .SetSceneFogParams 5, 52, -10.01
    .SetSceneFogColor 15724527
    .SetSceneFogEnable D_FALSE
 End With
 Set FR_Camera = D3D_Main.CreateFrame(FR_Root) 'Our Camera's Sub Frame. It goes under FR_Root in the 'Tree'.
For Genc = 1 To 4
   Set FR_Light(Genc) = D3D_Main.CreateFrame(FR_Root) 'Our Light's Sub Frame
Next Genc
  DS_Front.DrawText 200, 0, "Loading please wait .......", False  'Draw some text!
                                                                   ' sub frame.
 'That above code set up the hierarchy of frames where FR_Root is the parent, and the other frames are all
 'owned by it.

 FR_Root.SetSceneBackgroundRGB 0, 0, 0 'Set the background color. Use decimals, not the standerd 255 = max.
 CAMY = 0 ' height above floor                                                        'What I have here will make the background/sky 100% blue.
 FR_Camera.SetPosition Nothing, 112, CAMY + 3, 150 'Set The Camera position; X=1, Y=1, z=0.
 Set D3D_ViewPort = D3D_Main.CreateViewport(D3D_Device, FR_Camera, 0, 0, 640, 480) 'Make our viewport and set
                                                                                                                  'it our camera to be it.
 D3D_ViewPort.SetBack 400 'How far back it will draw the image. (Kinda like a visibility limit)
  ' set a few flags etc
  CAMROT = 0
  LightX = -55
  LightY = -4
  LightZ = 40
  LightPen = 2
  CAMROT = 0
  CollisionON = True
  LEVEl = 0
  liftUP = "stop"
  Shotfired = 0
  ShotVel = 0
  MeRotx = 0
  MeRotZ = 1
  DegRot = 0
  InfoOn = 0
  GunMove = -0.05
  GunCount = 0
  AMB = 0
  CollHIT = 16777215 ' white
  Foot = False

Timer5.Interval = 100

' this is your bullet
 Set MS_BULLET = D3D_Main.CreateMeshBuilder() 'Make the 3D Building Mesh
 Set FR_BULLET = D3D_Main.CreateFrame(FR_Root) 'Our Building (the 3D Thingy that will be placed in our world's)
 With MS_BULLET
     .LoadFromFile App.Path & "\models\MODEBULLE1.x", 0, 0, Nothing, Nothing
     .ScaleMesh 1, 1, 1 'Set the it's scale. This is used to make the object smaller or bigger. 1 makes
     .SetColor 235 '.SetTexture D3D_Main.LoadTexture(App.Path & "\textures\" & MODELTEX(Genc))
End With
FR_BULLET.AddVisual MS_BULLET
FR_BULLET.SetPosition Nothing, pos.x, pos.y, pos.z

' Gun

' this is your gun
 Set MS_GUN = D3D_Main.CreateMeshBuilder() 'Make the 3D Building Mesh
 Set FR_GUN = D3D_Main.CreateFrame(FR_Root) 'Our Building (the 3D Thingy that will be placed in our world's)
 With MS_GUN
     .LoadFromFile App.Path & "\models\MODEGUN002.x", 0, 0, Nothing, Nothing
     .ScaleMesh 1, 1, 1 'Set the it's scale. This is used to make the object smaller or bigger. 1 makes
     .SetTexture D3D_Main.LoadTexture(App.Path & "\textures\TEXTGUNB01.bmp")
End With
FR_GUN.AddVisual MS_GUN
FR_GUN.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z


LoadMap


FR_Building(5).GetPosition Nothing, RAMPXYZ1
FR_Building(6).GetPosition Nothing, RAMPXYZ2

Exit Sub
DXHELL:
MsgBox Err.Description
End Sub

'=====================================================================================
'                              MAIN LOOP
'=====================================================================================
Private Sub DX_Render()
On Error GoTo RENDERHELL:
 'Lets put our main loop. Make it loop until esc = true
 Do While esc = False
   On Local Error Resume Next 'Incase there is an error
   DoEvents 'Give the computer time to do what it needs to do.
   DX_Input 'Call the input sub.
   
   FpS = FpS + 1
   FR_Camera.GetPosition Nothing, pos
   D3D_ViewPort.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER 'Clear your viewport.
   D3D_Device.Update 'Update the Direct3D Device.
   D3D_ViewPort.Render FR_Root 'Render the 3D Objects (lights, and your building!)
   If InfoOn = 1 Then
    DS_Back.DrawText 10, 450, "Where am I ? X: " & pos.x & " Y: " & pos.y & " Z: " & pos.z & _
    " CAMROT " & CAMROT & " Collision " & CollHIT, False             'Draw some text!
    DS_Back.DrawText 10, 465, "Collision " & CollisionON & " number of objects " & Genc2 _
    & " Level:" & LEVEl & " fps " & oldfps & " AvFPS " & AvFPS & " shot " & Shotfired & " hit " & HIT & " sol " & Int(SOLvec.x) & " Bul " & Int(Bulletfire.x), False               'Draw some text!
    DS_Back.DrawText 10, 435, " merotx " & MeRotx & " merotz " & MeRotZ & " rotate " & DegRot, False
    DS_Back.DrawText 10, 405, "Keys - arrow keys forwards/back,left,right, I/O info on/off, Q/A look up/down, Z/X step sideways left/right", False
    DS_Back.DrawText 10, 420, "B/N open/close bayfloor doors, P transport to start position, ESC quit  " & " LX " & _
   LightX & " Ly " & LightY & " Lz " & LightZ & " Pen " & LightPen, False
   End If
   '& " LX " & _
   LightX & " Ly " & LightY & " Lz " & LightZ & " Pen " & LightPen
   DS_Front.Flip Nothing, DDFLIP_WAIT 'Flip the back buffer with the front buffer.
 Loop
 Exit Sub
RENDERHELL:
' MsgBox Err.Description
End Sub
'=====================================================================================
'                    END OF MAIN LOOP
'=====================================================================================


Private Sub DX_Input()
On Error GoTo DX_InputHELL:

    Const Sin5 = 8.715574E-02!  ' Sin(5°)
    Const Cos5 = 0.9961947!     ' Cos(5°)

FR_Building(22).GetPosition Nothing, AIRvec1
 DI_Device.GetDeviceStateKeyboard DI_State 'Get the array of keyboard keys and their current states
 FR_Camera.GetPosition Nothing, pos
 FR_GUN.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward


   
 If DI_State.Key(DIK_ESCAPE) <> 0 Then Call DX_Exit 'If user presses [esc] then exit end the program.
 
 If DI_State.Key(DIK_LEFT) <> 0 Then 'Quick Note: <> means 'does not'

     FR_GUN.SetOrientation FR_Camera, -Sin5, 0, Cos5, 0, 1, 0
     FR_BULLET.SetOrientation FR_GUN, -Sin5, 0, Cos5, 0, 1, 0
     FR_Camera.SetOrientation FR_Camera, -Sin5, 0, Cos5, 0, 1, 0
     DegRot = DegRot - 5
     DEGcheck
End If
 
 If DI_State.Key(DIK_RIGHT) <> 0 Then
    FR_GUN.SetOrientation FR_Camera, Sin5, 0, Cos5, 0, 1, 0
    FR_BULLET.SetOrientation FR_GUN, Sin5, 0, Cos5, 0, 1, 0
    FR_Camera.SetOrientation FR_Camera, Sin5, 0, Cos5, 0, 1, 0

    DegRot = DegRot + 5
    DEGcheck
 End If

 If (DI_State.Key(DIK_UP) <> 0) And (liftUP = "stop") Then
   NormCAM


    GUNanim
    FR_Camera.SetPosition FR_Camera, 0, 0, 1 'Move the viewport forward
    FR_Camera.GetPosition Nothing, pos
    FR_GUN.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward
    FR_BULLET.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward

  ' Timer3_Timer
   ' We're only going to check collision if collision flag is TRUE
   If CollisionON = True Then
        ColYMove

    If CollHIT <= 0 Then

        FR_Camera.SetPosition FR_Camera, 0, 0, -1 'Move the viewport back
        FR_Camera.GetPosition Nothing, pos
        FR_GUN.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward
        FR_BULLET.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward
    End If
    
   End If ' everything inside collision
 If Foot = False Then Foot = True
 End If

 If (DI_State.Key(DIK_DOWN)) <> 0 And (liftUP = "stop") Then
    NormCAM
    GUNanim
    FR_Camera.SetPosition FR_Camera, 0, 0, -1 'Move the viewport back
    FR_Camera.GetPosition Nothing, pos
    FR_GUN.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward
    FR_BULLET.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward
    ColYMove

    ' We're only going to check collision if collision flag is TRUE
    If CollisionON = True Then
     CollHIT = Form2.Picture1.Point(Int(pos.x), Int(pos.z))
    If CollHIT <= 0 Then
     FR_Camera.SetPosition FR_Camera, 0, 0, 1 'Move the viewport back
   FR_Camera.GetPosition Nothing, pos
   FR_GUN.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward
   FR_BULLET.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward
End If

        
        End If ' everything inside collision
   
   End If
 
 If DI_State.Key(DIK_Q) <> 0 Then
    If CAMROT < 50 Then
    PlaySoundWithPan 3, "Wall.wav", 70, 50    ' Play the buffer for this index
   ' Timer3_Timer
        CAMROT = CAMROT + 1
        'NormCAM
        FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 0.02, 0, 0, -0.02 'Move the viewport forward
    End If
 End If
 
 If DI_State.Key(DIK_A) <> 0 Then
 PlaySoundWithPan 3, "Wall.wav", 70, 50    ' Play the buffer for this index
    If CAMROT > -50 Then
   ' Timer3_Timer
        CAMROT = CAMROT - 1
        'NormCAM
        FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, -0.02, 0, 0, -0.02  'Move the viewport forward
    End If
 End If
 
 If DI_State.Key(DIK_C) <> 0 Then
        CollisionON = True
 End If
 
 If DI_State.Key(DIK_V) <> 0 Then
        CollisionON = False
 End If
 If DI_State.Key(DIK_B) <> 0 Then
        BAYdoor = "open"
 End If
  If DI_State.Key(DIK_N) <> 0 Then
        BAYdoor = "close"
 End If
 If DI_State.Key(DIK_I) <> 0 Then
        InfoOn = 1
 End If
  If DI_State.Key(DIK_O) <> 0 Then
        InfoOn = 0
 End If
 
 If DI_State.Key(DIK_SPACE) <> 0 Then
    If Shotfired = 0 Then

        FR_GUN.GetPosition Nothing, pos
        CameraPos = pos 'last good vector
         bullX = MeRotx
        bulletpos.y = pos.y
         bullZ = MeRotZ
        Shotfired = 1
        FireBULLET
    End If
 End If
 
 If DI_State.Key(DIK_M) <> 0 Then
    NormCAM
 End If
  If DI_State.Key(DIK_W) <> 0 Then
  FR_Camera.GetPosition Nothing, pos
  CAMY = pos.y
  CAMY = CAMY + 1
   FR_Camera.SetPosition Nothing, pos.x, CAMY, pos.z ' emergency teleport
 End If
   If DI_State.Key(DIK_S) <> 0 Then
  FR_Camera.GetPosition Nothing, pos
  CAMY = pos.y
  CAMY = CAMY - 1
   FR_Camera.SetPosition Nothing, pos.x, CAMY, pos.z ' emergency teleport
 End If
 If DI_State.Key(DIK_P) <> 0 Then
 NormCAM
 AMB = 0
 CAMY = 0
 PlaySoundWithPan 2, "Startsnd.wav", 70, 50    ' Play the buffer for this index
    FR_Camera.SetPosition Nothing, 150, CAMY + 3, 130 ' emergency teleport
 End If
 

 
 If (DI_State.Key(DIK_Z) <> 0) And (liftUP = "stop") Then
 'PlaySoundWithPan 3, "Wall.wav", 70, 50    ' Play the buffer for this index
   NormCAM
   GUNanim
   FR_Camera.SetPosition FR_Camera, -1, 0, 0 'Move the viewport left
   FR_Camera.GetPosition Nothing, pos
   FR_GUN.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward
   ColYMove

   ' We're only going to check collision if collision flag is TRUE
   If CollisionON = True Then
      CollHIT = Form2.Picture1.Point(Int(pos.x), Int(pos.z))
        If CollHIT <= 0 Then
            FR_Camera.SetPosition FR_Camera, 1, 0, 0 'Move the viewport back
            FR_Camera.GetPosition Nothing, pos
            FR_GUN.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward
            FR_BULLET.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward
        End If
    End If
 End If
 
  If (DI_State.Key(DIK_X) <> 0) And (liftUP = "stop") Then
    NormCAM
    GUNanim
    FR_Camera.SetPosition FR_Camera, 1, 0, 0 'Move the viewport left
    FR_Camera.GetPosition Nothing, pos
    FR_GUN.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward
    ColYMove
    ' We're only going to check collision if collision flag is TRUE
    If CollisionON = True Then
        CollHIT = Form2.Picture1.Point(Int(pos.x), Int(pos.z))
        If CollHIT <= 0 Then
            FR_Camera.SetPosition FR_Camera, -1, 0, 0 'Move the viewport back
            FR_Camera.GetPosition Nothing, pos
            FR_GUN.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward
            FR_BULLET.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z 'Move the viewport forward
        End If
  End If
 End If
 
 If Shotfired = 0 Then FR_BULLET.SetPosition Nothing, pos.x, pos.y, pos.z
 If pos.y >= 7 Then
    LEVEl = 1
 Else
    LEVEl = 0
 End If

 If LEVEl = 0 Then Form2.Picture1 = LoadPicture(App.Path & "\textures\" & "PLAN000001.bmp")
 If LEVEl = 1 Then Form2.Picture1 = LoadPicture(App.Path & "\textures\" & "PLAN000002.bmp")
 If CollHIT = 4144959 Then AIRlock1 = "open"
 If CollHIT = 16777215 Then AIRlock1 = "close"
 'Exit Sub
 If DI_State.Key(DIK_1) <> 0 Then
 LightX = LightX + 1
FR_Root.SetSceneFogParams LightX, LightY, 0.1
 End If
  If DI_State.Key(DIK_2) <> 0 Then
 LightX = LightX - 1
FR_Root.SetSceneFogParams LightX, LightY, 0.1
 End If
  If DI_State.Key(DIK_3) <> 0 Then
 LightY = LightY + 1
FR_Root.SetSceneFogParams LightX, LightY, 0.1
 End If
  If DI_State.Key(DIK_4) <> 0 Then
 LightY = LightY - 1
FR_Root.SetSceneFogParams LightX, LightY, 0.1
 End If
 Exit Sub
 
 '==================== we don't want to do this just yet

 If DI_State.Key(DIK_1) <> 0 Then
    LightX = LightX + 1
    FR_Light(2).SetPosition Nothing, LightX, LightY, LightZ
 End If
 If DI_State.Key(DIK_2) <> 0 Then
    LightX = LightX - 1
    FR_Light(2).SetPosition Nothing, LightX, LightY, LightZ
 End If
 If DI_State.Key(DIK_3) <> 0 Then
    LightY = LightY + 1
    FR_Light(2).SetPosition Nothing, LightX, LightY, LightZ
 End If
 If DI_State.Key(DIK_4) <> 0 Then
    LightY = LightY - 1
    FR_Light(2).SetPosition Nothing, LightX, LightY, LightZ
 End If
   If DI_State.Key(DIK_5) <> 0 Then
    LightZ = LightZ + 1
    FR_Light(2).SetPosition Nothing, LightX, LightY, LightZ
 End If
 If DI_State.Key(DIK_6) <> 0 Then
    LightZ = LightZ - 1
    FR_Light(2).SetPosition Nothing, LightX, LightY, LightZ
 End If
 
 
 Exit Sub
DX_InputHELL::
 MsgBox Err.Description

End Sub
Sub ColYMove()
   CollHIT = Form2.Picture1.Point(Int(pos.x), Int(pos.z))
        If CollHIT = 16777215 Then
        CAMY = 0
        ElseIf CollHIT = 15724527 Then
        CAMY = 1
        ElseIf CollHIT = 14671839 Then
        CAMY = 2
        ElseIf CollHIT = 13619151 Then
        CAMY = 3
        ElseIf CollHIT = 12566463 Then
        CAMY = 4
        ElseIf CollHIT = 1153775 Then
        CAMY = 5
        ElseIf CollHIT = 10461087 Then
        CAMY = 6
        ElseIf CollHIT = 9408399 Then
        CAMY = 7
        ElseIf CollHIT = 8355711 Then
        CAMY = 8
        ElseIf CollHIT = 7303025 Then
        CAMY = 9
        End If
                FR_Camera.GetPosition Nothing, pos
        FR_Camera.SetPosition Nothing, pos.x, CAMY + 3, pos.z
End Sub
Sub DEGcheck()
    If DegRot = 0 Then
        MeRotZ = 1
        MeRotx = 0
    End If
    If DegRot < 0 Then
        DegRot = 355
    End If
    If DegRot > 270 Then
        MeRotx = -(1 - ((DegRot - 270) / 90))
        MeRotZ = ((DegRot - 270) / 90)
    End If
    If DegRot = 270 Then
        MeRotx = -1
        MeRotZ = 0
    End If
    If (DegRot > 180) And (DegRot < 270) Then
        MeRotx = -((DegRot - 180) / 90)
        MeRotZ = -(1 - (DegRot - 180) / 90)
    End If
    If DegRot = 180 Then
        MeRotx = 0
        MeRotZ = -1
    End If
    If (DegRot > 90) And (DegRot < 180) Then
        MeRotx = (1 - (DegRot - 90) / 90)
        MeRotZ = -((DegRot - 90) / 90)
    End If
    If DegRot = 90 Then
        MeRotx = 1
        MeRotZ = 0
    End If
    If (DegRot > 0) And (DegRot < 90) Then
        MeRotx = ((DegRot - 0) / 90)
        MeRotZ = (1 - (DegRot - 0) / 90)
    End If
End Sub
Sub GUNanim()
   If GunCount > 10 Then GunCount = 0
   If GunCount < 4.9 Then GunMove = GunMove + 0.002
   If GunCount > 5 Then GunMove = GunMove - 0.002
   GunCount = GunCount + 1
End Sub
Private Sub NormCAM()
OLDCAMROT = CAMROT
        If CAMROT > 0 Then
            For Genc = OLDCAMROT To 1 Step -1
                FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, -0.02, 0, 0, -0.02  'Move the viewport forward
                CAMROT = CAMROT - 1
            Next Genc
        End If
        If CAMROT < 0 Then
            For Genc = CAMROT To -1 Step 1
                FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 0.02, 0, 0, -0.02  'Move the viewport forward
                CAMROT = CAMROT + 1
            Next Genc
        End If
End Sub

Private Sub FireBULLET()
    If Shotfired = 0 Then
        Shotfired = 1
    End If
End Sub
'=====================================================================================
Private Sub DX_Exit()
'Set DD_Main = Nothing
'Set DS_Front = Nothing
'Set DS_Back = Nothing
'Set D3D_Main = Nothing
'Set D3D_Device = Nothing
'Set FR_Root = Nothing
'Set FR_Camera = Nothing
'Set FR_Light = Nothing
'Set D3D_ViewPort = Nothing
'Set LT_Spot = Nothing
'Set LT_Ambient = Nothing
'For Genc = 1 To Genc2
'    Set MS_Building(Genc) = Nothing
'    Set FR_Building(Genc) = Nothing
'Next Genc
ShowCursor (1)
 Call DD_Main.RestoreDisplayMode
 Call DD_Main.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
 Call DI_Device.Unacquire

 'Restore all the devices

 End 'Ends the program.
End Sub
'=====================================================================================

Private Sub Form_Load()
HIT = 0
animFrame = 1
AIR1 = 0
AIR2 = 0
AIRlock1 = "closed"
AIRlock2 = "closed"
BAYdoor = "closed"
RAMPY = 0
AMBfix = 0
 Me.Show 'Some computers do weird stuff if you don't show the form.
 DoEvents 'Give the computer time to do what it needs to do
 DX_Init 'Initialize DirectX
 DX_MakeObjects 'Make frames, lights, and mesh(es)
 ' sound stuff ==========================================
 SetupDX7Sound Me              ' Assign DX7 to this Application
  SoundDir App.Path & "\Sound"  'Where is the applications sound stored

  CreateBuffers cSoundBuffers, "Default.wav" ' How many Channels/buffers (I used 10)
 PlaySoundWithPan 2, "Startsnd.wav", 70, 50    ' Play the buffer for this index
 DX_Render 'The Main Loop

End Sub




Private Sub Timer1_Timer()
'If Foot = True Then
     '  PlaySoundWithPan 1, "footsteps.wav", 75, 50    ' Play the buffer for this index
 '   Foot = False
'End If
End Sub

Private Sub Timer2_Timer()
oldfps = FpS
Totsecs = Totsecs + 1
If Totsecs > 0 Then
    totFrames = totFrames + FpS
    AvFPS = totFrames / Totsecs
End If
FpS = 0
'FR_Building(Genc2 - 7).SetRotation Nothing, 0, RotX, 0, RotX / 10 'Move the sphere
End Sub





Private Sub Timer3_Timer()

If (BAYdoor = "open") Then

If RAMPY <= 34 Then
RAMPY = RAMPY + 0.5
FR_Building(5).SetPosition Nothing, RAMPXYZ1.x + RAMPY, RAMPXYZ1.y, RAMPXYZ1.z
FR_Building(6).SetPosition Nothing, RAMPXYZ2.x + RAMPY, RAMPXYZ2.y, RAMPXYZ2.z

PlaySoundWithPan 9, "Reload.wav", 80, 50    ' Play the buffer for this index
End If
End If

If (BAYdoor = "close") Then

If RAMPY >= 0 Then
RAMPY = RAMPY - 0.5
FR_Building(5).SetPosition Nothing, RAMPXYZ1.x + RAMPY, RAMPXYZ1.y, RAMPXYZ1.z
FR_Building(6).SetPosition Nothing, RAMPXYZ2.x + RAMPY, RAMPXYZ2.y, RAMPXYZ2.z
PlaySoundWithPan 9, "Reload.wav", 80, 50     ' Play the buffer for this index
End If
End If
If RAMPY = 0 Then BAYdoor = "closed"
If RAMPY = 34 Then BAYdoor = "closed"
End Sub

Private Sub Timer4_Timer()

    If Shotfired = 1 Then
         
        
        ShotVel = ShotVel + 5
        If ShotVel = 5 Then
          PlaySoundWithPan 1, "Shoot.wav", 75, 50    ' Play the buffer for this index
          PlaySoundWithPan 9, "Reload.wav", 77, 50    ' Play the buffer for this index
        End If
        If ShotVel < 11 Then LT_Ambient.SetColorRGB 1.5, 0.9, 0.9
        
        
            FR_BULLET.SetPosition Nothing, CameraPos.x + (bullX * ShotVel), CameraPos.y - 0.1, CameraPos.z + (bullZ * ShotVel)
        If ShotVel < 30 Then
        
        GunMove = GunMove + 0.008
        End If
        If ShotVel > 30 Then GunMove = GunMove - 0.008
        FR_GUN.SetPosition Nothing, pos.x, pos.y + (GunMove - 0.75), pos.z
        If ShotVel > 60 Then
            ShotVel = 0
            Shotfired = 0
            GunMove = -0.05
        End If
        'FR_OBJECTS(1).GetPosition Nothing, SOLvec
        'FR_BULLET.GetPosition Nothing, Bulletfire
        'If Int(SOLvec.X) = Int(Bulletfire.X) Then
        'HIT = 1
        'End If
    End If
End Sub

Private Sub Timer5_Timer()

If (AMB < 10) And (AMB > -1) Then
    AMB = AMB + 0.5

        LT_Ambient.SetColorRGB 1.2 - (0.1 * AMB), 1.2 - (0.1 * AMB), 1.2 - (0.1 * AMB)

Else

     LT_Ambient.SetColorRGB 0.7 + (LEVEl / 10), 0.7 + (LEVEl / 10), 0.7 + (LEVEl / 10)

      AMB = -1
          
End If


End Sub
Private Sub Timer6_Timer()
'23 = baylock 02
If (AIRlock1 = "open") Then

If AIR1 <= 29 Then
AIR1 = AIR1 + 0.5
FR_Building(23).SetPosition Nothing, AIRvec1.x, AIRvec1.y, AIRvec1.z + AIR1


PlaySoundWithPan 9, "Reload.wav", 80, 50    ' Play the buffer for this index
End If
End If
If (AIRlock1 = "close") Then

If AIR1 >= 0 Then
AIR1 = AIR1 - 0.5
FR_Building(23).SetPosition Nothing, AIRvec1.x, AIRvec1.y, AIRvec1.z + AIR1


PlaySoundWithPan 7, "Reload.wav", 80, 50    ' Play the buffer for this index
End If
End If
End Sub

Private Sub Timer7_Timer()
'If animFrame = 1 Then
'FR_Building(47).GetPosition Nothing, ENvec
'FR_Building(47).SetPosition Nothing, ENvec.X, ENvec.Y + 500, ENvec.z
'FR_Building(48).SetPosition Nothing, ENvec.X, ENvec.Y, ENvec.z
'animFrame = 2
'Else
'FR_Building(48).GetPosition Nothing, ENvec
'FR_Building(48).SetPosition Nothing, ENvec.X, ENvec.Y + 500, ENvec.z
'FR_Building(47).SetPosition Nothing, ENvec.X, ENvec.Y, ENvec.z
' animFrame = 1
'End If
End Sub
 Sub LoadMap()
'FR_Light(2).SetPosition Nothing, -5, -5, -60  'Set our 'point' light.
' Set LT_Spot(2) = D3D_Main.CreateLightRGB(D3DRMLIGHT_SPOT, 7, 0.9, 0.9)  'Set the light type and it's color.
'With LT_Spot(2)
'  .SetPenumbra LightPen
'End With
' FR_Light(2).AddLight LT_Spot(2) 'Add the light to it's frame.

FR_Light(2).SetPosition Nothing, LightX - 50, LightY, LightZ + 50 'Set our 'point' light.
 Set LT_Spot(2) = D3D_Main.CreateLightRGB(D3DRMLIGHT_POINT, 0.5, 0.5, 0.5)  'Set the light type and it's color.
 FR_Light(2).AddLight LT_Spot(2) 'Add the light to it's frame.
 
'  FR_Light(4).SetPosition Nothing, LightX, LightY, LightZ + 50  'Set our 'point' light.
' Set LT_Spot(4) = D3D_Main.CreateLightRGB(D3DRMLIGHT_POINT, 0.5, 0.5, 0.5)  'Set the light type and it's color.
' FR_Light(4).AddLight LT_Spot(2) 'Add the light to it's frame.

 Set LT_Ambient = D3D_Main.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.2, 0.2, 0.2)  'Create our ambient light.
 FR_Root.AddLight LT_Ambient 'Add the ambient light to the root frame.


Set MS_BACKOB(1) = D3D_Main.CreateMeshBuilder() 'Make the 3D Building Mesh
Set FR_BACKOB(1) = D3D_Main.CreateFrame(FR_Root) 'Our BACKOB (the 3D Thingy that will be placed in our world's)

With MS_BACKOB(1)
     .LoadFromFile App.Path & "\models\Space01.x", 0, 0, Nothing, Nothing
     .ScaleMesh 1, 1, 1 'Set the it's scale. This is used to make the object smaller or bigger. 1 makes
     .SetTexture D3D_Main.LoadTexture(App.Path & "\textures\starfield1.bmp")
End With

With FR_BACKOB(1)
  .AddVisual MS_BACKOB(1)
  '   .SetPosition Nothing, 50, -0.5, 50 ' we'll leave the position as the one set in TrueSpace
End With

For Genc3 = 1 To 4
    LoadMesh Genc3, "wall0" & Genc3 & ".x", "TEXTRUST01.bmp"
Next Genc3

LoadMesh 5, "retractFloor01.x", "TEXTRUST01.bmp"

LoadMesh 6, "retractFloor02.x", "TEXTRUST01.bmp"

LoadMesh 7, "dropship01.x", "TEXTRUST01.bmp"

LoadMesh 8, "dropship02.x", "TEXTRUST01.bmp"

LoadMesh 9, "bayfloor01.X", "TEXTRUST02.bmp"
 
LoadMesh 10, "bayramp.X", "TEXTRUST01.bmp"

'LoadMesh 11, "walltest.X", "TEXTRUST01.bmp"

LoadMesh 12, "wall04WIN.X", "TEXTRUST11.bmp"

LoadMesh 13, "bayback.X", "TEXTRUST11.bmp"
 
LoadMesh 14, "bayroof.X", "TEXTRUST11.bmp"

LoadMesh 15, "wall02WIN.X", "TEXTRUST11.bmp"

LoadMesh 16, "ramprail01.X", "TEXTRUST03.bmp"

LoadMesh 17, "ramprail02.X", "TEXTRUST03.bmp"

LoadMesh 18, "level1rail1.x", "TEXTRUST03.bmp"

LoadMesh 19, "level1rail3d.x", "TEXTRUST03.bmp"

LoadMesh 20, "level1rail3u.x", "TEXTRUST03.bmp"

LoadMesh 21, "level1rail2.x", "TEXTRUST03.bmp"

LoadMesh 22, "bayLock001.x", "TEXTRUST01.bmp"

LoadMesh 23, "bayLock002.x", "TEXTRUST03.bmp"

LoadMesh 24, "bayback2.x", "TEXTRUST11.bmp"

LoadMesh 25, "rearbayfloor.x", "TEXTRUST01.bmp"

LoadMesh 26, "rearbaywall.x", "TEXTRUST11.bmp"

LoadMesh 27, "forefloor.x", "TEXTRUST02.bmp"

LoadMesh 28, "wall12WIN.x", "TEXTRUST11.bmp"

LoadMesh 29, "wall22WIN.x", "TEXTRUST11.bmp"

LoadMesh 30, "engine01.x", "TEXTRUST01.bmp"

LoadMesh 31, "engine02.x", "TEXTRUST01.bmp"

LoadMesh 32, "wing01.x", "TEXTRUST01.bmp"

LoadMesh 33, "wing02.x", "TEXTRUST01.bmp"

LoadMesh 34, "nose01.x", "TEXTRUST01.bmp"

LoadMesh 35, "nose02.x", "TEXTRUST11.bmp"

LoadMesh 36, "forebridge.x", "TEXTRUST01.bmp"

LoadMesh 37, "baybridge.x", "TEXTRUST01.bmp"

LoadMesh 38, "bridgerail01.x", "TEXTRUST03.bmp"

LoadMesh 39, "bridgerail02.x", "TEXTRUST03.bmp"

LoadMesh 40, "level1rail11.x", "TEXTRUST03.bmp"

LoadMesh 41, "level1rail12.x", "TEXTRUST03.bmp"

LoadMesh 42, "level1rail13.x", "TEXTRUST03.bmp"

LoadMesh 43, "level1rail23.x", "TEXTRUST03.bmp"

'LoadMesh 44, "level1rail14.x", "TEXTRUST03.bmp"

'LoadMesh 45, "level1rail15.x", "TEXTRUST03.bmp"

'LoadMesh 46, "level1rail24.x", "TEXTRUST03.bmp"

LoadMesh 47, "bayfloor02.X", "TEXTRUST02.bmp"

LoadMesh 48, "upperfloor01.X", "TEXTRUST01.bmp"


End Sub



Sub LoadMesh(MeshNo, MeshName, Optional MeshTexture As String, Optional XPOS, Optional YPOS, Optional ZPOS)
Form1.Label1.Caption = "Loading...............100% Complete" '" & MeshNo
Set MS_Building(1) = D3D_Main.CreateMeshBuilder() 'Make the 3D Building Mesh
Set FR_Building(MeshNo) = D3D_Main.CreateFrame(FR_Root) 'Our Building (the 3D Thingy that will be placed in our world's)

With MS_Building(1)
     .LoadFromFile App.Path & "\models\" & MeshName, 0, 0, Nothing, Nothing
     .ScaleMesh 1, 1, 1 'Set the it's scale. This is used to make the object smaller or bigger. 1 makes
     If MeshTexture <> "" Then
        .SetTexture D3D_Main.LoadTexture(App.Path & "\textures\" & MeshTexture)
     End If
End With

With FR_Building(MeshNo)
  .AddVisual MS_Building(1)
  'if xpos <>"" then
  '   .SetPosition Nothing, 50, -0.5, 50 ' we'll leave the position as the one set in TrueSpace
End With

End Sub
