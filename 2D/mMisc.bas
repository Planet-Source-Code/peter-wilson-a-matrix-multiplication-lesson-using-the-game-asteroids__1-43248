Attribute VB_Name = "mMisc"
Option Explicit


Private Const g_sngPIDivideBy180 = 0.0174533!

Public Sub DrawCrossHairs()

    frmCanvas.DrawWidth = 1
    frmCanvas.ForeColor = RGB(32, 32, 32)
    
    ' Draw vertical line.
    frmCanvas.Line (frmCanvas.Width / 2, 0)-(frmCanvas.Width / 2, frmCanvas.Height)

    ' Draw horizontal line.
    frmCanvas.Line (0, frmCanvas.Height / 2)-(frmCanvas.Width, frmCanvas.Height / 2)
    
End Sub
Public Function ConvertDeg2Rad(Degress As Single) As Single

    ' Converts Degrees to Radians
    ConvertDeg2Rad = Degress * (g_sngPIDivideBy180)
    
End Function

Public Sub Debug_PrintMatrix(m1 As mdrMATRIX3x3)

    With m1
        
        frmCanvas.ForeColor = RGB(192, 192, 192)
        frmCanvas.CurrentX = 0
        frmCanvas.CurrentY = 0
        
        frmCanvas.Print .rc11, .rc12, .rc13
        frmCanvas.Print .rc21, .rc22, .rc23
        frmCanvas.Print .rc31, .rc32, .rc33
        
    End With
    
End Sub

Public Function GetRNDNumberBetween(Min As Variant, Max As Variant) As Single

    GetRNDNumberBetween = (Rnd * (Max - Min)) + Min

End Function

Public Function zTest()

''    Dim sngMax As Single
''    Dim sngMin As Single
''    Dim sngRND As Single
''
''    sngMax = -10000
''    sngMin = 10000
''
''    Dim intN
''    For intN = 0 To 1000
''        sngRND = GetRNDNumberBetween(12.5, 87.6)
''        If sngRND > sngMax Then sngMax = sngRND
''        If sngRND < sngMin Then sngMin = sngRND
''    Next intN
''
''    Debug.Print Format(sngMin, "0.0000"), sngMax
    
End Function
