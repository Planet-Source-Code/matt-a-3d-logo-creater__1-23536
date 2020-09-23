VERSION 5.00
Begin VB.Form frmBkgCol 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Background Color"
   ClientHeight    =   2340
   ClientLeft      =   8820
   ClientTop       =   3615
   ClientWidth     =   1620
   ControlBox      =   0   'False
   Icon            =   "frmBackCol.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   1620
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
   Begin VB.VScrollBar vscCol 
      Height          =   1575
      Index           =   2
      LargeChange     =   5
      Left            =   720
      Max             =   100
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.VScrollBar vscCol 
      Height          =   1575
      Index           =   1
      LargeChange     =   5
      Left            =   360
      Max             =   100
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   945
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Width           =   975
   End
   Begin VB.VScrollBar vscCol 
      Height          =   1575
      Index           =   0
      LargeChange     =   5
      Left            =   0
      Max             =   100
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmBkgCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub cmdOk_Click()
   Me.Hide
End Sub

Private Sub Form_Load()
   picColor.BackColor = RGB(backRGB.X, backRGB.Y, backRGB.Z)
   vscCol(0).Value = backRGB.X * 100
   vscCol(1).Value = backRGB.Y * 100
   vscCol(2).Value = backRGB.Z * 100

End Sub

Private Sub vscCol_Change(Index As Integer)
   Select Case Index
      Case 0
         backRGB.X = vscCol(Index).Value / 100
      Case 1
         backRGB.Y = vscCol(Index).Value / 100
      Case 2
         backRGB.Z = vscCol(Index).Value / 100
   End Select
   glClearColor backRGB.X, backRGB.Y, backRGB.Z, 0#         ' Set the background color
   picColor.BackColor = RGB(backRGB.X * 255, backRGB.Y * 255, backRGB.Z * 255)  'refresh the preview picture
End Sub
