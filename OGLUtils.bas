Attribute VB_Name = "OpenGL"
Option Explicit
'see http://nehe.gamedev.net/opengl.asp
'for great tutorials on opengl
'tutorials are for c++ but many include
'code for VB and most of the c++ opengl code
'is easily modified into opengl VB code
'good portion of this module code taken and
'altered from lesson14  VB code by Ross Dawson
Public logoString As String
Public frm As Form
Public quadratic As GLUquadricObj               ' Storage For Our Quadratic Objects
Private Declare Function EnumDisplaySettings Lib "USER32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "USER32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Private Declare Function CreateIC Lib "GDI32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Long) As Long
Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Type DEVMODE
    dmDeviceName        As String * CCDEVICENAME
    dmSpecVersion       As Integer
    dmDriverVersion     As Integer
    dmSize              As Integer
    dmDriverExtra       As Integer
    dmFields            As Long
    dmOrientation       As Integer
    dmPaperSize         As Integer
    dmPaperLength       As Integer
    dmPaperWidth        As Integer
    dmScale             As Integer
    dmCopies            As Integer
    dmDefaultSource     As Integer
    dmPrintQuality      As Integer
    dmColor             As Integer
    dmDuplex            As Integer
    dmYResolution       As Integer
    dmTTOption          As Integer
    dmCollate           As Integer
    dmFormName          As String * CCFORMNAME
    dmUnusedPadding     As Integer
    dmBitsPerPel        As Integer
    dmPelsWidth         As Long
    dmPelsHeight        As Long
    dmDisplayFlags      As Long
    dmDisplayFrequency  As Long
End Type
Private hrc As Long
Public base As GLuint         ' Base Display List For The Font Set
Public Type glRPos   'used for rot,pos,logoRGB,backRGB
   X As GLfloat      'Instead of having 12 variables have 4 types
   Y As GLfloat
   Z As GLfloat
End Type
Public rot As glRPos 'Holds the rotation values for logo
Public pos As glRPos 'Holds the position of the logo
Public LogoRGB As glRPos 'Holds the red(x) green(y) blue(z) values 0 to 1
Public backRGB As glRPos 'holds colors r(x)g(y)b(z) values 0 to 1
Public gmf(256) As GLYPHMETRICSFLOAT   ' Storage For Information About Our Font
Public FontDepth As Double 'actually how thick the font is
Private OldWidth As Long   'Used to store info on previous screen
Private OldHeight As Long  'so can be returned properly
Private OldBits As Long
Private OldVertRefresh As Long
Public Texture(2) As GLuint ' Storage for the texture with the 3 different filters 0,1,2
Public TextBool As Boolean 'Set to true on successful load of bitmap texture
Public TextureFilename As String 'Stores the filename before calling the loadgltextures 'Should have rewritten it to pass directly to it and avoided the global string
Public filter As GLuint  ' Which Filter To Use 0 1 or 2 still not sure what gluint really is integer would probably work just as well?
Public light As Boolean    ' Lighting ON / OFF
Public fontFamily As String 'Holds the font name
Public FescVal As Integer  'Holds the font escapement value
Public glSlant As GLdouble
Public glAspect As GLdouble 'window aspect value
Public fontCharset As Boolean 'If true use symbols instead of letters. If the current font family doesn't have characters then a default one will be used.
                              'if a symbol set is selected then must be set of the proper symbols may not be used
Private Function LoadBMP(ByVal Filename As String, ByRef Texture() As GLuint, ByRef Height As Long, ByRef Width As Long) As Boolean
   ' Open a file.
   ' The file should be BMP with pictures 8x8,16x16,32x32,64x64,128x128,256x256 ...You get the point
   ' You get no notice of failure other than no texture is applied Most likely
   ' failure reason is bitmap dimensions must be as stated above
   frmMain.picTexture.Picture = LoadPicture(Filename)
   If LTSVerify = False Then
      MsgBox "Textures must be perfect squares (ie.. 8*8, 16*16 .... 512*512)", vbOKOnly, "Texture failed to load"
   End If

   CreateTextureMapFromImage frmMain.picTexture, Texture(), Height, Width
   LoadBMP = True
End Function

Public Function LoadGLTextures() As Boolean
   ' This is the function called to load texture
   ' Set TextureFilename before calling
   ' if succesfull then returns true
   Dim Status As Boolean
   Dim h As Long
   Dim w As Long
   Dim TextureImage() As GLbyte
   Status = False                        ' Status Indicator
   If LoadBMP(TextureFilename, TextureImage(), h, w) Then
      ' Load The Bitmap, Check For Errors, If Bitmap's Not Found Quit
      Status = True                          ' Set The Status To TRUE
      glGenTextures 2, Texture(0)            ' Create The Textures
      ' Create Nearest Filtered Texture
      glBindTexture glTexture2D, Texture(0)
      glTexParameteri glTexture2D, tpnTextureMagFilter, GL_NEAREST '( NEW )
      glTexParameteri glTexture2D, tpnTextureMinFilter, GL_NEAREST '( NEW )
      glTexImage2D glTexture2D, 0, 3, w, h, 0, GL_RGB, GL_UNSIGNED_BYTE, TextureImage(0, 0, 0)
      ' Create Linear Filtered Texture
      glBindTexture glTexture2D, Texture(1)
      glTexParameteri glTexture2D, tpnTextureMinFilter, GL_LINEAR     ' Linear Filtering
      glTexParameteri glTexture2D, tpnTextureMagFilter, GL_LINEAR     ' Linear Filtering
      glTexImage2D glTexture2D, 0, 3, w, h, 0, GL_RGB, GL_UNSIGNED_BYTE, TextureImage(0, 0, 0)
      ' Create MipMapped Texture
      glBindTexture glTexture2D, Texture(2)
      glTexParameteri glTexture2D, tpnTextureMagFilter, GL_LINEAR
      glTexParameteri glTexture2D, tpnTextureMinFilter, GL_LINEAR_MIPMAP_NEAREST
      gluBuild2DMipmaps glTexture2D, 3, w, h, GL_RGB, GL_UNSIGNED_BYTE, VarPtr(TextureImage(0, 0, 0))
      ' Texturing Contour Anchored To The Object
      glTexGeni tcS, tgTextureGenMode, tgmObjectLinear
      ' Texturing Contour Anchored To The Object
      glTexGeni tcT, tgTextureGenMode, tgmObjectLinear
      glEnable glcTextureGenS          ' Auto Texture Generation
      glEnable glcTextureGenT          ' Auto Texture Generation
   End If
   Erase TextureImage   ' Free the texture image memory
   LoadGLTextures = Status 'set the return value true if success or false on failure
End Function
Private Sub CreateTextureMapFromImage(pict As PictureBox, ByRef TextureImg() As GLbyte, ByRef Height As Long, ByRef Width As Long)
    ' Create the array as needed for the image.
    frm.MousePointer = vbHourglass 'This can take a bit
    pict.ScaleMode = 3                  ' Pixels
    Height = pict.ScaleHeight
    Width = pict.ScaleWidth
    ReDim TextureImg(2, Height - 1, Width - 1)
    
    ' Fill the array with the bitmap data...  This could take
    ' a while...
    
    Dim X As Long, Y As Long
    Dim c As Long
    Dim yloc As Long
    For X = 0 To Width - 1
        For Y = 0 To Height - 1
            c = pict.Point(X, Y)                ' Returns in long format.
            yloc = Height - Y - 1
            TextureImg(0, X, yloc) = c And 255
            TextureImg(1, X, yloc) = (c And 65280) \ 256
            TextureImg(2, X, yloc) = (c And 16711680) \ 65536
        Next Y
    Next X
    frm.MousePointer = vbDefault
End Sub


Public Function DrawGLScene() As Boolean
' Here's Where We Do All The Drawing
   glClear clrColorBufferBit Or clrDepthBufferBit  ' Color The Screen And The Depth Buffer
   glLoadIdentity                                  ' Reset The Current Modelview Matrix
   If TextBool = True Then
      glEnable glcTexture2D
      glBindTexture GL_TEXTURE_2D, Texture(filter) ' Select Our Texture
   End If
   If TextBool = False Then
      glDisable glcTexture2D
   End If
   glTranslatef pos.X, pos.Y, pos.Z                 ' Move to current x,y,z position
   glRotatef rot.X, 1#, 0#, 0#                 ' Rotate On The X Axis
   glRotatef rot.Y, 0#, 1#, 0#            ' Rotate On The Y Axis
   glRotatef rot.Z, 0#, 0#, 1#           ' Rotate On The Z Axis
   'Colors Based On The RGB values provided by user
   glColor3f LogoRGB.X, LogoRGB.Y, LogoRGB.Z 'logo values range from 0 to 1 but nothing bad happens if you go over 1
   glPrint (logoString)         ' Print GL Text To The Screen
   glPopMatrix
   DrawGLScene = True                              ' Everything Went OK
End Function
Public Sub BuildFont(frm As Form)                    ' Build Our Bitmap Font
   'SYMBOL_CHARSET=2 ANSI_CHARSET = 0
   'for later when I rework this down a bit smaller
   KillFont 'Just to make sure previous font's cleared out when rebuilding fonts
   frm.MousePointer = vbHourglass 'takes a sec show the hourglass
   Dim hfont As Long                       ' Windows Font ID
   base = glGenLists(256)                   ' Storage For 256 Characters
   If fontCharset = False Then    'Dont seem to notice much difference by altering
                                  'the fontwieght width and height so stopped messing with it
      hfont = CreateFont(-12, 0, FescVal, 0, FW_BOLD, False, False, False, _
               ANSI_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, _
               FF_DONTCARE Or DEFAULT_PITCH, fontFamily)
   ElseIf fontCharset = True Then
      hfont = CreateFont(-12, 0, FescVal, 0, FW_BOLD, False, False, False, _
               SYMBOL_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, _
               FF_DONTCARE Or DEFAULT_PITCH, fontFamily)
   End If
   SelectObject frm.hDC, hfont                ' Selects The Font just created
   wglUseFontOutlines frm.hDC, 0, 255, base, 0#, FontDepth, WGL_FONT_POLYGONS, gmf(0)
   frm.MousePointer = vbDefault
End Sub
Private Sub KillFont()                     ' Delete The Font
   glDeleteLists base, 256                ' Delete All 256 Characters
End Sub
Public Sub glPrint(ByVal s As String)                ' Custom GL "Print" Routine
    ' we are just going to provide a simple print routine just like normal basic
    Dim b() As Byte
    Dim i As Integer
    Dim length As Double
    If Len(s) > 0 Then              ' only if the pass a string
        ReDim b(Len(s))             ' array of bytes to hold the string
        For i = 1 To Len(s)         ' for each character
            b(i - 1) = Asc(Mid$(s, i, 1)) ' convert from unicode to ascii
            length = length + gmf(b(i)).gmfCellIncX      ' Increase Length By Each Characters Width
        Next
        b(Len(s)) = 0               ' null terminated
        glTranslatef -length / 2, 0#, 0#          ' Center Our Text On The Screen
        glPushAttrib amListBit               ' Pushes The Display List Bits
        glListBase base                  ' Sets The Base Character to 32
        glCallLists Len(s), GL_UNSIGNED_BYTE, b(0)   ' Draws The Display List Text
        glPopAttrib                      ' Pops The Display List Bits
    End If
End Sub
Public Sub ReSizeGLScene(ByVal Width As GLsizei, ByVal Height As GLsizei)
 ' Resize And Initialize The GL Window
    If Height = 0 Then              ' Prevent A Divide By Zero By
        Height = 1                  ' Making Height Equal One
    End If
    glViewport 0, 0, Width, Height  ' Reset The Current Viewport
    glMatrixMode mmProjection       ' Select The Projection Matrix
    glLoadIdentity                  ' Reset The Projection Matrix
    ' Calculate The Aspect Ratio Of The Window
    gluPerspective glSlant, glAspect, 0.1, 200#
    glMatrixMode mmModelView         ' Select The Modelview Matrix
    glLoadIdentity                  ' Reset The Modelview Matrix
End Sub
Public Function InitGL() As Boolean
   ' All Setup For OpenGL Goes Here
   glEnable glcTexture2D
   glShadeModel smSmooth               ' Enables Smooth Shading
   glClearColor 0#, 1#, 1#, 0.5       ' green blue Background to start 'last 0 is intensity which I don't use in this program
   glClearDepth 1#                     ' Depth Buffer Setup
   glEnable glcDepthTest               ' Enables Depth Testing
   glDepthFunc cfLEqual                ' The Type Of Depth Test To Do
   glHint htPerspectiveCorrectionHint, hmNicest    ' Really Nice Perspective Calculations
   glEnable glcLight0                     ' Enable Default Light (Quick And Dirty)   ( NEW )
   glEnable glcColorMaterial              ' Enable Coloring Of Material          ( NEW )
   InitGL = True                       ' Initialization Went OK
End Function
Public Sub KillGLWindow()
' Properly Kill The Window
   If hrc Then                                     ' Do We Have A Rendering Context?
      If wglMakeCurrent(0, 0) = 0 Then             ' Are We Able To Release The DC And RC Contexts?
         MsgBox "Release Of DC And RC Failed.", vbInformation, "SHUTDOWN ERROR"
      End If
      If wglDeleteContext(hrc) = 0 Then           ' Are We Able To Delete The RC?
         MsgBox "Release Rendering Context Failed.", vbInformation, "SHUTDOWN ERROR"
      End If
      hrc = 0                                     ' Set RC To NULL
   End If
   KillFont                     ' Destroy The Font
    ' Note
    ' The form owns the device context (hDC) window handle (hWnd) and class (RTThundermain)
    ' so we do not have to do all the extra work

End Sub

Private Sub SaveCurrentScreen()
    ' Save the current screen resolution, bits, and Vertical refresh
    Dim ret As Long
    ret = CreateIC("DISPLAY", "", "", 0&)
    OldWidth = GetDeviceCaps(ret, HORZRES)
    OldHeight = GetDeviceCaps(ret, VERTRES)
    OldBits = GetDeviceCaps(ret, BITSPIXEL)
    OldVertRefresh = GetDeviceCaps(ret, VREFRESH)
    ret = DeleteDC(ret)
End Sub

Private Function FindDEVMODE(ByVal Width As Integer, ByVal Height As Integer, ByVal Bits As Integer, Optional ByVal VertRefresh As Long = -1) As DEVMODE
    ' locate a DEVMOVE that matches the passed parameters
    Dim ret As Boolean
    Dim i As Long
    Dim dm As DEVMODE
    i = 0
    Do  ' enumerate the display settings until we find the one we want
        ret = EnumDisplaySettings(0&, i, dm)
        If dm.dmPelsWidth = Width And _
            dm.dmPelsHeight = Height And _
            dm.dmBitsPerPel = Bits And _
            ((dm.dmDisplayFrequency = VertRefresh) Or (VertRefresh = -1)) Then Exit Do ' exit when we have a match
        i = i + 1
    Loop Until (ret = False)
    FindDEVMODE = dm
End Function
Private Sub ResetDisplayMode()
   Dim dm As DEVMODE             ' Device Mode
   dm = FindDEVMODE(OldWidth, OldHeight, OldBits, OldVertRefresh)
   dm.dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
   If OldVertRefresh <> -1 Then
      dm.dmFields = dm.dmFields Or DM_DISPLAYFREQUENCY
   End If
   ' Try To Set Selected Mode And Get Results.  NOTE: CDS_FULLSCREEN Gets Rid Of Start Bar.
   If (ChangeDisplaySettings(dm, CDS_FULLSCREEN) <> DISP_CHANGE_SUCCESSFUL) Then
    
   ' If The Mode Fails, Offer Two Options.  Quit Or Run In A Window.
      MsgBox "The Requested Mode Is Not Supported By Your Video Card", , "NeHe GL"
   End If

End Sub

Private Sub SetDisplayMode(ByVal Width As Integer, ByVal Height As Integer, ByVal Bits As Integer, ByRef fullscreen As Boolean, Optional VertRefresh As Long = -1)
    Dim dmScreenSettings As DEVMODE             ' Device Mode
    Dim p As Long
    SaveCurrentScreen                           ' save the current screen attributes so we can go back later
    dmScreenSettings = FindDEVMODE(Width, Height, Bits, VertRefresh)
    dmScreenSettings.dmBitsPerPel = Bits
    dmScreenSettings.dmPelsWidth = Width
    dmScreenSettings.dmPelsHeight = Height
    dmScreenSettings.dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
    If VertRefresh <> -1 Then
        dmScreenSettings.dmDisplayFrequency = VertRefresh
        dmScreenSettings.dmFields = dmScreenSettings.dmFields Or DM_DISPLAYFREQUENCY
    End If
    ' Try To Set Selected Mode And Get Results.  NOTE: CDS_FULLSCREEN Gets Rid Of Start Bar.
    If (ChangeDisplaySettings(dmScreenSettings, CDS_FULLSCREEN) <> DISP_CHANGE_SUCCESSFUL) Then
    
        ' If The Mode Fails, Offer Two Options.  Quit Or Run In A Window.
        If (MsgBox("The Requested Mode Is Not Supported By" & vbCr & "Your Video Card. Use Windowed Mode Instead?", vbYesNo + vbExclamation, "NeHe GL") = vbYes) Then
            fullscreen = False                  ' Select Windowed Mode (Fullscreen=FALSE)
        Else
            ' Pop Up A Message Box Letting User Know The Program Is Closing.
            MsgBox "Program Will Now Close.", vbCritical, "ERROR"
            End                   ' Exit And Return FALSE
        End If
    End If
End Sub

Public Function CreateGLWindow(frm As Form, Width As Integer, Height As Integer, Bits As Integer, fullscreenflag As Boolean) As Boolean
    Dim PixelFormat As GLuint                       ' Holds The Results After Searching For A Match
    Dim pfd As PIXELFORMATDESCRIPTOR                ' pfd Tells Windows How We Want Things To Be
    pfd.cAccumAlphaBits = 0
    pfd.cAccumBits = 0
    pfd.cAccumBlueBits = 0
    pfd.cAccumGreenBits = 0
    pfd.cAccumRedBits = 0
    pfd.cAlphaBits = 0
    pfd.cAlphaShift = 0
    pfd.cAuxBuffers = 0
    pfd.cBlueBits = 0
    pfd.cBlueShift = 0
    pfd.cColorBits = Bits
    pfd.cDepthBits = 24
    pfd.cGreenBits = 0
    pfd.cGreenShift = 0
    pfd.cRedBits = 0
    pfd.cRedShift = 0
    pfd.cStencilBits = 0
    pfd.dwDamageMask = 0
    pfd.dwflags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
    pfd.dwLayerMask = 0
    pfd.dwVisibleMask = 0
    pfd.iLayerType = PFD_MAIN_PLANE
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    
    PixelFormat = ChoosePixelFormat(frm.hDC, pfd)
    If PixelFormat = 0 Then                     ' Did Windows Find A Matching Pixel Format?
        KillGLWindow                            ' Reset The Display
        MsgBox "Can't Find A Suitable PixelFormat.", vbExclamation, "ERROR"
        CreateGLWindow = False                  ' Return FALSE
    End If

    If SetPixelFormat(frm.hDC, PixelFormat, pfd) = 0 Then ' Are We Able To Set The Pixel Format?
        KillGLWindow                            ' Reset The Display
        MsgBox "Can't Set The PixelFormat.", vbExclamation, "ERROR"
        CreateGLWindow = False                           ' Return FALSE
    End If
    
    hrc = wglCreateContext(frm.hDC)
    If (hrc = 0) Then                           ' Are We Able To Get A Rendering Context?
        KillGLWindow                            ' Reset The Display
        MsgBox "Can't Create A GL Rendering Context.", vbExclamation, "ERROR"
        CreateGLWindow = False                  ' Return FALSE
    End If

    If wglMakeCurrent(frm.hDC, hrc) = 0 Then    ' Try To Activate The Rendering Context
        KillGLWindow                            ' Reset The Display
        MsgBox "Can't Activate The GL Rendering Context.", vbExclamation, "ERROR"
        CreateGLWindow = False                  ' Return FALSE
    End If
    frm.Show                                    ' Show The Window
    SetForegroundWindow frm.hWnd                ' Slightly Higher Priority
    frm.SetFocus                                ' Sets Keyboard Focus To The Window
    ReSizeGLScene frm.ScaleWidth, frm.ScaleHeight ' Set Up Our Perspective GL Screen

    If Not InitGL() Then                        ' Initialize Our Newly Created GL Window
        KillGLWindow                            ' Reset The Display
        MsgBox "Initialization Failed.", vbExclamation, "ERROR"
        CreateGLWindow = False                   ' Return FALSE
    End If

    
    CreateGLWindow = True                       ' Success

End Function

Sub Main()
   logoString = "3D Logo Creator"
   pos.X = 0
   pos.Y = 0
   pos.Z = -3
   rot.X = 340
   rot.Y = 0
   rot.Z = 0
   FescVal = 0
   FontDepth = 0.7
   glAspect = 3.2
   glSlant = 45
   'set the logo and background colors
   LogoRGB.X = 1: LogoRGB.Y = 1: LogoRGB.Z = 0
   backRGB.X = 0: backRGB.Y = 1: backRGB.Z = 1
   Dim Done As Boolean
   Done = False
   fontFamily = "Comic Sans MS"
   ' Create Our OpenGL Window
   Set frm = New frmMain 'from now on form1 will be referred to as frm
   If Not CreateGLWindow(frm, 800, 600, 24, False) Then 'if you want full screen set fullscreen to true
      Done = True                             ' Quit If Window Was Not Created
   End If
   glEnable glcLighting 'default  is lighting on
   BuildFont frm
   Do While Not Done
      ' Draw The Scene.  Watch Quit Messages From DrawGLScene()
      If Not DrawGLScene Then                 ' Updating View Only If Active
         Unload frm                          ' DrawGLScene Signalled A Quit
      Else                                    ' Not Time To Quit, Update Screen
         SwapBuffers (frm.hDC)               ' Swap Buffers (Double Buffering)
         DoEvents
      End If
      Done = frm.Visible = False              ' if the form is not visible then we are done
   Loop
   ' Shutdown
   Set frm = Nothing
   End
End Sub

Private Function LTSVerify() As Boolean
   'Insure that the texture being loaded is square and
   'between 8*8 .... and 512*512 'could go larger still as long as its double
   'but don't think there is much need for textures above 512
   LTSVerify = False
   Select Case frmMain.picTexture.Width
      Case 8
         If frmMain.picTexture.Width = frmMain.picTexture.Height Then LTSVerify = True
      Case 16
         If frmMain.picTexture.Width = frmMain.picTexture.Height Then LTSVerify = True
      Case 32
         If frmMain.picTexture.Width = frmMain.picTexture.Height Then LTSVerify = True
      Case 64
         If frmMain.picTexture.Width = frmMain.picTexture.Height Then LTSVerify = True
      Case 128
         If frmMain.picTexture.Width = frmMain.picTexture.Height Then LTSVerify = True
      Case 256
         If frmMain.picTexture.Width = frmMain.picTexture.Height Then LTSVerify = True
      Case 512
         If frmMain.picTexture.Width = frmMain.picTexture.Height Then LTSVerify = True
   End Select
End Function






