VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Gear Toy"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   Icon            =   "Form1.frx":0000
   ScaleHeight     =   606
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   577
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar sbTweak 
      Height          =   255
      LargeChange     =   5
      Left            =   4320
      Max             =   45
      TabIndex        =   14
      Top             =   8640
      Width           =   1815
   End
   Begin VB.HScrollBar sbRotate 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   360
      TabIndex        =   13
      Top             =   8640
      Width           =   3855
   End
   Begin VB.HScrollBar sbAdjust 
      Height          =   255
      Index           =   3
      LargeChange     =   20
      Left            =   240
      Max             =   100
      Min             =   5
      TabIndex        =   11
      Top             =   8040
      Value           =   10
      Width           =   3855
   End
   Begin VB.HScrollBar sbAdjust 
      Height          =   255
      Index           =   2
      LargeChange     =   20
      Left            =   240
      Max             =   100
      Min             =   5
      TabIndex        =   9
      Top             =   7440
      Value           =   10
      Width           =   3855
   End
   Begin VB.HScrollBar sbAdjust 
      Height          =   255
      Index           =   1
      LargeChange     =   20
      Left            =   240
      Max             =   200
      Min             =   5
      TabIndex        =   7
      Top             =   6840
      Value           =   77
      Width           =   3855
   End
   Begin VB.HScrollBar sbAdjust 
      Height          =   255
      Index           =   0
      LargeChange     =   20
      Left            =   240
      Max             =   200
      Min             =   5
      TabIndex        =   5
      Top             =   6240
      Value           =   31
      Width           =   3855
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFC0&
      Height          =   5280
      Left            =   120
      ScaleHeight     =   348
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   556
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   8400
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Tweak"
      Height          =   195
      Index           =   5
      Left            =   4320
      TabIndex        =   17
      Top             =   8400
      Width           =   495
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Rotate"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   8400
      Width           =   480
   End
   Begin VB.Label lblSelection 
      BackColor       =   &H000000FF&
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   5640
      Width           =   3900
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "ToothDepth"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Bore"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   7200
      Width           =   330
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Diameter"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Teeth"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   6000
      Width           =   540
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   195
      Index           =   3
      Left            =   3840
      TabIndex        =   4
      Top             =   7800
      Width           =   45
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   195
      Index           =   2
      Left            =   3840
      TabIndex        =   3
      Top             =   7200
      Width           =   45
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   195
      Index           =   1
      Left            =   3840
      TabIndex        =   2
      Top             =   6600
      Width           =   45
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   195
      Index           =   0
      Left            =   3840
      TabIndex        =   1
      Top             =   6000
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gear Toy by Paul Bahlawan
'
'Built on Gear-Box by Emilio P.G. Ficara http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=73622&lngWId=1


Option Explicit

Dim myGears(2) As GEAR
Dim userSelect As Long

Private Sub Display()
   
    myGears(0).rAngle = sbRotate.Value * PI / 180
    myGears(1).rAngle = (360 - sbRotate.Value * PI / 180) * (myGears(0).Teeth / myGears(1).Teeth) + sbTweak.Value * PI / 180
    myGears(2).rAngle = (sbRotate.Value * PI / 180) * (myGears(0).Teeth / myGears(2).Teeth)
   
    picDisplay.Cls
    DrawGear myGears(0), picDisplay
    DrawGear myGears(1), picDisplay
    DrawGear myGears(2), picDisplay

End Sub

Private Sub Form_Load()
    
    'Set the gears with some initial values
    With myGears(0)
        .cX = 120
        .cY = 175
        .bRad = sbAdjust(2).Value
        .pRad = sbAdjust(1).Value
        .Teeth = sbAdjust(0).Value
        .tDepth = sbAdjust(3).Value
        .Colour = vbRed
    End With
    myGears(1) = myGears(0)
    myGears(2) = myGears(0)
    MakeCompatible myGears(1), myGears(0), 13
    MakeCompatible myGears(2), myGears(0), 43
    With myGears(1)
        .cX = myGears(0).cX + CenterDistance(myGears(0), myGears(1))
        .cY = 175
        .Colour = vbGreen
    End With
    With myGears(2)
        .cX = myGears(1).cX + CenterDistance(myGears(1), myGears(2))
        .cY = 175
        .Colour = vbBlue
    End With

    Display
End Sub

Private Sub Form_Resize()
    picDisplay.Width = Form1.ScaleWidth - 20
    Display
End Sub

Private Sub sbRotate_Change()
    If sbRotate.Value = 360 Then sbRotate.Value = 0
    Display
End Sub

Private Sub sbRotate_Scroll()
    Display
End Sub

Private Sub sbTweak_Change()
    Display
End Sub

Private Sub sbTweak_Scroll()
    Display
End Sub

Private Sub picDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    For i = 0 To 2 'User selects a gear to edit
        If Abs(myGears(i).cX - X) < 10 And Abs(myGears(i).cY - Y) < 10 Then
            userSelect = -1
            sbAdjust(0).Value = myGears(i).Teeth
            sbAdjust(2).Value = myGears(i).bRad
            sbAdjust(3).Value = myGears(i).tDepth
            lblSelection.BackColor = myGears(i).Colour
            If i > 0 Then
                sbAdjust(1).Enabled = False
                sbAdjust(3).Enabled = False
                lblInfo(1).Caption = Int(myGears(i).pRad)
            Else
                sbAdjust(1).Enabled = True
                sbAdjust(3).Enabled = True
                sbAdjust(1).Value = myGears(i).pRad
                lblInfo(1).Caption = Int(myGears(i).pRad)
            End If
            userSelect = i
            Exit For
        End If
    Next i
End Sub

Private Sub sbAdjust_Change(Index As Integer)
    lblInfo(Index).Caption = sbAdjust(Index).Value
    If userSelect < 0 Then Exit Sub
    UpdateGear
    Display
End Sub

Private Sub sbAdjust_Scroll(Index As Integer)
    lblInfo(Index).Caption = sbAdjust(Index).Value
    UpdateGear
    Display
End Sub

Private Sub UpdateGear()
    With myGears(userSelect)
        .bRad = sbAdjust(2).Value
        .pRad = sbAdjust(1).Value
        .Teeth = sbAdjust(0).Value
        .tDepth = sbAdjust(3).Value
    End With
    
    MakeCompatible myGears(1), myGears(0), myGears(1).Teeth
    MakeCompatible myGears(2), myGears(0), myGears(2).Teeth
    
    myGears(1).cX = myGears(0).cX + CenterDistance(myGears(0), myGears(1))
    myGears(2).cX = myGears(1).cX + CenterDistance(myGears(1), myGears(2))
End Sub
