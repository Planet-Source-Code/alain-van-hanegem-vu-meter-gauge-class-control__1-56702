VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VU meter/gauge test"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   579
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox p1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Index           =   2
      Left            =   5280
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   215
      TabIndex        =   2
      Top             =   0
      Width           =   3255
   End
   Begin VB.PictureBox p1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Index           =   1
      Left            =   2040
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   215
      TabIndex        =   1
      Top             =   0
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   1440
   End
   Begin VB.PictureBox p1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   0
      Left            =   0
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "right button = direct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "left button = smooth (gauge control needle setting)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   2400
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Click on the controls to move the needle to the mouse cursor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "A FULLY customizable VU meter/gauge class control made from scratch! FAST GDI!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' V1.0.0 - 2004-10-12 by Alain van Hanegem
' E-mail: alain@decos.nl

Option Explicit

' Demonstration of VU/gauge control: 3 types are created below. I put 2 default in the
' control for easy starting! ;)

' Speedo gauge is not fed with random numbers, the VU meters are for nice animation

Dim c As Integer

Private Sub Form_Activate()
    Dim cx As Long
    Dim cy As Long
    
    cx = p1(0).ScaleWidth \ 2
    cy = p1(0).ScaleHeight \ 2
    v(0).Init_Picture p1(0).hDC, p1(0).Image, p1(0).ScaleWidth, p1(0).ScaleHeight
    v(0).SetVUDefaults 0, cx, cy
    v(0).Draw

    cx = p1(1).ScaleWidth \ 2
    cy = p1(1).ScaleHeight \ 2 + 35
    v(1).Init_Picture p1(1).hDC, p1(1).Image, p1(1).ScaleWidth, p1(1).ScaleHeight
    v(1).SetVUDefaults 1, cx, cy
    v(1).Draw

    cx = p1(2).ScaleWidth \ 2
    cy = p1(2).ScaleHeight \ 2 + 35
    v(2).Init_Picture p1(2).hDC, p1(2).Image, p1(2).ScaleWidth, p1(2).ScaleHeight
    v(2).SetVUDefaults 1, cx, cy
    v(2).Draw

    p1(0).Refresh
    p1(1).Refresh
    p1(2).Refresh
    
    Randomize

    Timer1.enabled = True
    Timer1.interval = 20        ' 50 Hz interval is nice (almost no CPU usage). change it to 10 (100 Hz) is even nicer!
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shutdown
End Sub

' mouse click on the control feed back: left button = set value normaly; right button = set immediately

Private Sub p1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim r As Double
    
    r = v(Index).VU_MouseClick(x, y)
    If Button = 1 Then
        v(Index).SetNeedleValue r
    End If
    
    If Button = 2 Then
        v(Index).SetNeedleValueDirect r
    End If
End Sub

' timer event for calling the AnimationLoop of each VU control (which makes it move smoooooth)

Private Sub Timer1_Timer()
    v(0).AnimationLoop
    v(0).Draw
    p1(0).Refresh

    v(1).AnimationLoop
    v(1).Draw
    p1(1).Refresh

    v(2).AnimationLoop
    v(2).Draw
    p1(2).Refresh
    
    If Rnd < 0.1 Then
        v(1).SetNeedleValue (Rnd * 60) - 40
        v(2).SetNeedleValue (Rnd * 60) - 40
    End If
End Sub

