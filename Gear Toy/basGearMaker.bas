Attribute VB_Name = "basGearMaker"
Option Explicit

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type GEAR
    cX As Long      'location X
    cY As Long      'location Y
    bRad As Single  'bore radius (center hole)
    pRad As Single  'pitch radius
    Teeth As Long   'number of teeth
    tDepth As Single 'tooth depth
    rAngle As Single 'rotational angle (in radians)
    Colour As Long  'aka color
End Type

Public Const PI As Single = 3.14159265358979

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Sub DrawGear(ByRef dGear As GEAR, Destination As PictureBox)
    Dim r1 As Single
    Dim r2 As Single
    Dim da As Single
    Dim Angle As Single
    Dim i As Long
    Dim PolySet() As POINTAPI

    With dGear
        ReDim PolySet(3 + .Teeth * 4)
        
        r1 = .pRad - .tDepth / 2 'radius to tooth root
        r2 = .pRad + .tDepth / 2 'radius to tooth tip
        da = (2 * PI) / .Teeth / 4 'one fourth pitch angle
        
        'Make gear
        ReDim PolySet(3 + .Teeth * 4)
        
        For i = 0 To .Teeth
            Angle = i * 2 * PI / .Teeth + .rAngle
            PolySet(0 + i * 4).X = r1 * Cos(Angle) + .cX
            PolySet(0 + i * 4).Y = r1 * Sin(Angle) + .cY
            
            PolySet(1 + i * 4).X = r2 * Cos(Angle + da) + .cX
            PolySet(1 + i * 4).Y = r2 * Sin(Angle + da) + .cY
            
            PolySet(2 + i * 4).X = r2 * Cos(Angle + 2 * da) + .cX
            PolySet(2 + i * 4).Y = r2 * Sin(Angle + 2 * da) + .cY
            
            PolySet(3 + i * 4).X = r1 * Cos(Angle + 3 * da) + .cX
            PolySet(3 + i * 4).Y = r1 * Sin(Angle + 3 * da) + .cY
        Next i
        
        Destination.ForeColor = vbYellow
        Destination.FillColor = .Colour
        Polygon Destination.hdc, PolySet(0), .Teeth * 4
        
        'Make center hole
        Destination.FillColor = Destination.BackColor
        Destination.Circle (.cX, .cY), .bRad, vbYellow
    End With

End Sub

Public Sub MakeCompatible(Destination As GEAR, Source As GEAR, newTeeth As Long)
    Dim p As Single
    
    p = (PI * Source.pRad * 2) / Source.Teeth
    Destination.pRad = (p * newTeeth / PI) / 2
    Destination.Teeth = newTeeth
    Destination.tDepth = Source.tDepth
    
End Sub

Public Function CenterDistance(A As GEAR, B As GEAR) As Long

    CenterDistance = 2 + (A.Teeth + B.Teeth) / 2 / (A.Teeth / A.pRad / 2)

End Function
