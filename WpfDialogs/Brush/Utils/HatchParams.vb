
Public Class HatchParams
    Public Property HatchStyle As HatchStyle
    Public Property Background As Brush
    Public Property Foreground As Brush
    Public Property PenThickness As Double
    Public Property HatchTransform As Transform
    Public Property ShapeTransform As Transform

    Public Sub New()
        HatchStyle = HatchStyle.Horizontal
        Background = Brushes.White
        Foreground = Brushes.Black
        PenThickness = 1.0
    End Sub

    Public Sub New(HatchStyle As WpfDialogs.HatchStyle, Background As Brush, Foreground As Brush, PenThickness As Double, HatchTransform As Transform, ShapeTransform As Transform)
        Me.HatchStyle = HatchStyle
        Me.Background = Background
        Me.Foreground = Foreground
        Me.PenThickness = PenThickness
        Me.HatchTransform = HatchTransform
        Me.ShapeTransform = ShapeTransform
    End Sub

End Class

Public Enum HatchStyle
    Horizontal
    Vertical
    ForwardDiagonal
    BackwardDiagonal
    LargeGrid
    DiagonalCross
    Percent05
    Percent10
    Percent20
    Percent25
    Percent30
    Percent40
    Percent50
    Percent60
    Percent70
    Percent75
    Percent80
    Percent90
    LightDownwardDiagonal
    LightUpwardDiagonal
    DarkDownwardDiagonal
    DarkUpwardDiagonal
    WideDownwardDiagonal
    WideUpwardDiagonal
    LightVertical
    LightHorizontal
    NarrowVertical
    NarrowHorizontal
    DarkVertical
    DarkHorizontal
    DashedDownwardDiagonal
    DashedUpwardDiagonal
    DashedHorizontal
    DashedVertical
    SmallConfetti
    LargeConfetti
    Zigzag
    Wave
    DiagonalBrick
    HorizontalBrick
    Weave
    Plaid
    Divot
    DottedGrid
    DottedDiamond
    Shingle
    Trellis
    Sphere
    SmallGrid
    SmallCheckerboard
    LargeCheckerboard
    OutlinedDiamond
    SolidDiamond
    TwoOverlappingEllipses
    DotFill
End Enum

