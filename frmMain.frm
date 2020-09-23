VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   " Logo Creator"
   ClientHeight    =   2685
   ClientLeft      =   2595
   ClientTop       =   3465
   ClientWidth     =   6825
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   455
   Begin VB.PictureBox picTexture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picStore 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComDlg.CommonDialog dlgLogo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Times New Roman"
      Min             =   8
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuText 
         Caption         =   "&New Logo"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save Logo &As"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Logo"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuColors 
      Caption         =   "&Effects"
      Begin VB.Menu mnuLogoColor 
         Caption         =   "&Logo Color"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuBackgroundColor 
         Caption         =   "&Background Color"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnucolSep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFonts 
         Caption         =   "&Fonts"
         Begin VB.Menu mnuEscapement 
            Caption         =   "&Escapement"
            Begin VB.Menu mnuIncEsc 
               Caption         =   "&Increase Escapement"
               Shortcut        =   ^I
            End
            Begin VB.Menu mnuDecEsc 
               Caption         =   "&Decrease Escapement"
               Shortcut        =   ^O
            End
         End
         Begin VB.Menu mnuFamily 
            Caption         =   "&Family"
            Begin VB.Menu mnuComic 
               Caption         =   "Comic Sans MS"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuVerdana 
               Caption         =   "Verdana"
            End
            Begin VB.Menu mnuTimesNRoman 
               Caption         =   "Times New Roman"
            End
            Begin VB.Menu mnuArial 
               Caption         =   "Arial"
            End
            Begin VB.Menu mnuCustom 
               Caption         =   "&Custom"
               Shortcut        =   %{BKSP}
            End
         End
         Begin VB.Menu mnuSymbol 
            Caption         =   "Enable &Symbols"
         End
      End
      Begin VB.Menu mnuLoadTexture 
         Caption         =   "Load Logo &Texture"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuDelTexture 
         Caption         =   "&Delete Logo Texture"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuTextFilter 
         Caption         =   "&Texture Filter"
         Begin VB.Menu mnuNearFilter 
            Caption         =   "Nearest Filter"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuLinearFilter 
            Caption         =   "Linear Filter"
         End
         Begin VB.Menu mnuMipfilter 
            Caption         =   "MipMapped Filter"
         End
      End
      Begin VB.Menu mnuSepLight 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuLighting 
         Caption         =   "&Lighting"
         Begin VB.Menu mnuEnableLight 
            Caption         =   "Enabled"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuDisableLight 
            Caption         =   "Disabled"
         End
      End
   End
   Begin VB.Menu mnuHA 
      Caption         =   "&Help / About"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Alot more global variable
'than I like to use but
'wrote this a little to quick to worry
'frmMain scale mode set to pixels
'frmMain.AutoRedraw must be false
Dim logoFileName As String 'used so after first save can use just save with no dialogue



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 36 'home key
         'return position and angles to start up defaults
         pos.X = 0
         pos.Y = 0
         pos.Z = -3
         rot.X = 340 'a bit of an x angle looks better than no angle at all
         rot.Y = 0
         rot.Z = 0
         FontDepth = 0.7 'set fontdepth back
         FescVal = 0 'set font escapement value back
         glSlant = 45#
         glAspect = 3.2
         ReSizeGLScene ScaleWidth, ScaleHeight 'Resize with glSlant back to default 45
         BuildFont frm 'Rebuild the font with the defualt depth to 0.7
      Case 37 'arrow left
         pos.X = pos.X - 0.1
      Case 38 'arrow up
         pos.Y = pos.Y + 0.1
      Case 39 'arrow right
         pos.X = pos.X + 0.1
      Case 40 'arrow down
         pos.Y = pos.Y - 0.1
      Case 33 'page up
         pos.Z = pos.Z - 0.1
      Case 34 'page down
         pos.Z = pos.Z + 0.1
      Case 45 'Insert key for glSlant
         glSlant = glSlant - 0.1
         If glSlant < 360 Then glSlant = glSlant + 360
         ReSizeGLScene ScaleWidth, ScaleHeight
      Case 46 'Delete key for glSlant
         glSlant = glSlant + 0.1
         If glSlant < 360 Then glSlant = glSlant - 360
         ReSizeGLScene ScaleWidth, ScaleHeight
      Case 97 'keypad #1
         rot.Z = rot.Z - 0.5
         If rot.Z < 0 Then rot.Z = rot.Z + 360
      Case 98  'keypad #2
         rot.X = rot.X + 0.5
         If rot.X > 360 Then rot.X = rot.X - 360
      Case 99 'keypad #3 for glAspect
         glAspect = glAspect - 0.1
         ReSizeGLScene ScaleWidth, ScaleHeight 'must call resize to put it in effect
      Case 100 'keypad #4
         rot.Y = rot.Y - 0.5
         If rot.Y < 0 Then rot.Y = rot.Y + 360
      Case 101 'keypad #5 similar to home key
         rot.X = 340
         rot.Y = 0
         rot.Z = 0
         glSlant = 45#
         glAspect = 3.2
         ReSizeGLScene ScaleWidth, ScaleHeight 'Resize with defaults
      Case 102 'keypad #6
         rot.Y = rot.Y + 0.5
         If rot.Y > 360 Then rot.Y = rot.Y - 360
      Case 103 'keypad #7 for glaspect
         glAspect = glAspect + 0.1
         ReSizeGLScene ScaleWidth, ScaleHeight 'must call resize to put it in effect
      Case 104 'keypad #8
         rot.X = rot.X - 0.5
         If rot.X < 0 Then rot.X = rot.X + 360
      Case 105 'keypad #9
         rot.Z = rot.Z + 0.5
         If rot.Z > 360 Then rot.Z = rot.Z - 360
      Case 187 '+ key
         FontDepth = FontDepth + 0.1 'Increase the depth a little bit (it only takes a small amount)
         'I wont check for a max here I've tested it quite high and havn't run into trouble yet :)
         frm.MousePointer = vbHourglass 'takes a sec show the hourglass
         BuildFont frm 'call the font creator with the new font depth
         frm.MousePointer = vbDefault 'finished return to default pointer
      Case 189 '- Key
         FontDepth = FontDepth - 0.1 'Decrease the depth a little bit (it only takes a small amount)
         If FontDepth < 0.1 Then FontDepth = 0.1 'Keep it from going less than .1
         frm.MousePointer = vbHourglass 'takes a sec show the hourglass
         BuildFont frm 'call the font creator with the new font depth
         frm.MousePointer = vbDefault 'finished return to default pointer
   End Select
End Sub
Private Sub Form_Load()
   'Initialize menu
   mnuDelTexture.Enabled = False
   mnuTextFilter.Enabled = False
End Sub

Private Sub Form_Resize()
   'resizing the form requires resizing the glscene
   On Error GoTo errorHandler
   ReSizeGLScene ScaleWidth, ScaleHeight
   Exit Sub
errorHandler:
   Exit Sub
End Sub
Private Sub Form_Unload(Cancel As Integer)
   'when exiting the program the GLWindow must be killed
   KillGLWindow
End Sub
Private Sub mnuAbout_Click()
'Not much of an about at this time
MsgBox "Made by Matt ", vbOKOnly, "About 3D Logo Creator"

End Sub
Private Sub mnuArial_Click()
   If fontFamily = "Arial" Then Exit Sub
   fontFamily = "Arial"
   mnuArial.Checked = True
   mnuVerdana.Checked = False
   mnuComic.Checked = False
   mnuTimesNRoman.Checked = False
   mnuCustom.Checked = False
   BuildFont frm
End Sub



Private Sub mnuBackgroundColor_Click()
   'The common dialogue didn't work well for this
   'the color values for rgb need to be between
   '0 and 1 so made a custom color picker
   frmBkgCol.Show
End Sub
Private Sub mnuComic_Click()
   If fontFamily = "Comic Sans MS" Then Exit Sub
   fontFamily = "Comic Sans MS"
   mnuVerdana.Checked = False
   mnuComic.Checked = True
   mnuTimesNRoman.Checked = False
   mnuArial.Checked = False
   mnuCustom.Checked = False
   BuildFont frm
End Sub

Private Sub mnuCustom_Click()
   'Note Somethings wrong with my vb
   'the common dialogue refuses to show fonts
   'so just had to do it the wrong way instead
   fontFamily = InputBox("Enter the name of a font to try", "Custom Font", "Symbol")
   mnuVerdana.Checked = False
   mnuComic.Checked = False
   mnuTimesNRoman.Checked = False
   mnuArial.Checked = False
   mnuCustom.Checked = True
   BuildFont frm
End Sub

Private Sub mnuDecEsc_Click()
   FescVal = FescVal - 10
   If FescVal < 0 Then FescVal = FescVal + 3600
   BuildFont frm
End Sub

Private Sub mnuDelTexture_Click()
   TextBool = False
   mnuDelTexture.Enabled = False
   mnuTextFilter.Enabled = False
End Sub

Private Sub mnuDisableLight_Click()
   'Set check to this menu item
   'and disable lighting
   'Most logos look bad without lighting
   'but certain textures look intresting without
   'so I added it
   mnuEnableLight.Checked = False
   mnuDisableLight.Checked = True
   glDisable glcLighting
End Sub

Private Sub mnuEnableLight_Click()
   'Set check to this menu item
   'and enable lighting
   'Most logos look bad without lighting
   'but certain textures look intresting without
   'so I added it
   mnuEnableLight.Checked = True
   mnuDisableLight.Checked = False
   glEnable glcLighting
   
End Sub

Private Sub mnuExit_Click()
   KillGLWindow
   End
End Sub
Private Sub mnuHelp_Click()
   frmHelp.Show
End Sub





Private Sub mnuIncEsc_Click()
   FescVal = FescVal + 10
   If FescVal > 3600 Then FescVal = FescVal - 3600
   BuildFont frm
End Sub

Private Sub mnuLinearFilter_Click()
   'A better filter
   mnuNearFilter.Checked = False
   mnuLinearFilter.Checked = True
   mnuMipfilter.Checked = False
   filter = 1
End Sub
Private Sub mnuLoadTexture_Click()
   On Error GoTo errorHandler
   dlgLogo.InitDir = App.Path & "\textures"
   dlgLogo.filter = "bitmap texture|*.bmp|jpeg texture|*.jpg|gif texture|*.gif"
   dlgLogo.DialogTitle = "Load Bitmap Texture"
   dlgLogo.ShowOpen
   dlgLogo.FilterIndex = 1
   TextureFilename = dlgLogo.Filename
   If LoadGLTextures Then 'Once Global String TextureFilename is set load it up
         TextBool = True 'if successful then set to true
   End If
   If TextBool = True Then
      mnuDelTexture.Enabled = True
      mnuTextFilter.Enabled = True
   End If
   Exit Sub
errorHandler:
   frm.MousePointer = vbDefault 'in case loadgltextures failed
   TextBool = False 'stick with colors
   Exit Sub
End Sub
Private Sub mnuLogoColor_Click()
   'The common dialogue didn't work well for this
   'the color values for rgb need to be between
   '0 and 1 so made a custom color picker
   frmLogoCol.Show
End Sub
Private Sub mnuMipfilter_Click()
   'The best Filter but again depends
   'on the bitmap
   mnuNearFilter.Checked = False
   mnuLinearFilter.Checked = False
   mnuMipfilter.Checked = True
   filter = 2
End Sub
Private Sub mnuNearFilter_Click()
   'The lowest quality filter
   'But some textures actually look better
   'try them all out
   mnuNearFilter.Checked = True
   mnuLinearFilter.Checked = False
   mnuMipfilter.Checked = False
   filter = 0
End Sub


Private Sub mnuSave_Click()
   'High failure rate on saving
   'The menu grays out a section of
   'the image. Once you save through the menu
   'Check image. If it didn't work right
   'Try saving again with control S so nothing is over the form
   On Error GoTo ErrHandler
   If logoFileName = "" Then
      frm.SetFocus 'The form must have full focus to get a clean picture
      Call DrawGLScene 'Redraw the screen to make sure its clean
      DoEvents 'Make sure its caught up
      Set picStore.Picture = CaptureClient(frm) 'Capture the forms picture into a hidden picture box
      dlgLogo.CancelError = True
      dlgLogo.filter = "bitmap|*.bmp"  'Would like to find out how to save in other formats beside bmp
      dlgLogo.DialogTitle = "Save 3D Logo"
      dlgLogo.ShowSave
      DoEvents
      logoFileName = dlgLogo.Filename
      SavePicture picStore.Picture, logoFileName
      Exit Sub
   Else
      frm.SetFocus 'The form must have full focus to get a clean picture
      Call DrawGLScene 'Redraw the screen to make sure its clean
      DoEvents 'Make sure its caught up
      Set picStore.Picture = CaptureClient(frm) 'Capture the forms picture into a hidden picture box
      DoEvents
      SavePicture picStore.Picture, logoFileName
      Exit Sub
   End If
ErrHandler:
   Exit Sub
End Sub
Private Sub mnuSaveAs_Click()
   'High failure rate on saving
   'The menu grays out a section of
   'the image. Once you save through the menu
   'Check image. If it didn't work right
   'Try saving again with control S so nothing is over the form
   On Error GoTo ErrHandler
   frm.SetFocus 'The form must have full focus to get a clean picture
   Call DrawGLScene 'Redraw the screen to make sure its clean
   DoEvents 'Make sure its caught up
   Set picStore.Picture = CaptureClient(frm) 'Capture the forms picture into a hidden picture box
   dlgLogo.CancelError = True
   dlgLogo.filter = "bitmap|*.bmp"  'Would like to find out how to save in other formats beside bmp
   dlgLogo.DialogTitle = "Save 3D Logo"
   dlgLogo.ShowSave
   DoEvents
   logoFileName = dlgLogo.Filename
   SavePicture picStore.Picture, logoFileName
   Exit Sub
ErrHandler:
   Exit Sub
End Sub
Private Sub mnuSymbol_Click()
   'If you want symbols or be able to use webdings
   'the buildfont must know with fontcharset as true
   'Perhaps someday I'll figure out how to make text
   'and symbols :)
   mnuSymbol.Checked = Not mnuSymbol.Checked
   fontCharset = Not fontCharset
   BuildFont frm
End Sub
Private Sub mnuText_Click()
   Dim OldLogoString As String
   OldLogoString = logoString
   'logoString stores the string being displayed
   'Switching it will auto change on the programs next loop
   logoString = InputBox("Enter Logo Text", "3D Logo Creator", logoString)
   If Len(logoString) > 256 Then 'max string size 256 characters
      MsgBox "Maximum length of string 256 characters", vbOKOnly
      logoString = OldLogoString 'if string to long revert to previous string
   End If
   If logoString = "" Then logoString = OldLogoString
   'return position and angles to default
   pos.X = 0
   pos.Y = 0
   pos.Z = -3 'pos.Z should always be negative or wont be visible
   rot.X = 340
   rot.Y = 0
   rot.Z = 0
   FescVal = 0
   FontDepth = 0.7
   glAspect = 3.2
   glSlant = 45
End Sub
Private Sub mnuTimesNRoman_Click()
   If fontFamily = "Times New Roman" Then Exit Sub
   fontFamily = "Times New Roman"
   mnuVerdana.Checked = False
   mnuComic.Checked = False
   mnuTimesNRoman.Checked = True
   mnuArial.Checked = False
   mnuCustom.Checked = False
   BuildFont frm
End Sub
Private Sub mnuVerdana_Click()
   If fontFamily = "Verdana" Then Exit Sub
   fontFamily = "Verdana"
   mnuVerdana.Checked = True
   mnuComic.Checked = False
   mnuTimesNRoman.Checked = False
   mnuArial.Checked = False
   mnuCustom.Checked = False
   BuildFont frm
End Sub
