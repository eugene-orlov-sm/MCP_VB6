Attribute VB_Name = "MTrial"
Option Explicit

Private total As Long
Private evaluations As Long


Public Sub Main()
    total = -999
    evaluations = 0
    Outer 7
End Sub


Private Sub Outer(ByVal seed As Long)
    Dim adjusted As Long
    adjusted = seed + 5
    Inner adjusted, adjusted
End Sub


Private Sub Inner(ByVal inputValue As Long, ByRef copy As Long, Optional ByVal label As String = "ready")
    Dim total As Long
    Dim emptyText As String
    Dim emptyVariant As Variant
    Dim fraction As Double
    Dim flag As Boolean
    total = inputValue * 3
    fraction = 1.25
    flag = True
    Touch total
    Debug.Print "Trial result=" & CStr(total)
End Sub


Private Sub Touch(ByRef target As Long)
    target = target + 1
End Sub


Private Function Probe() As Long
    evaluations = evaluations + 1
    Probe = evaluations
End Function
