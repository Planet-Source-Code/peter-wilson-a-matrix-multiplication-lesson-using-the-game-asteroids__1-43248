Attribute VB_Name = "mKeyboard"
Option Explicit

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub ProcessKeyboardInput()

    Static s_strPreviousValue As String
    Static s_blnKeyDeBounce As Boolean
    
    Dim lngKeyState As Long
    
    lngKeyState = GetKeyState(vbKeyLeft)
    If (lngKeyState And &H8000) Then m_Enemies(0).WorldX = m_Enemies(0).WorldX - 1000
    
    lngKeyState = GetKeyState(vbKeyRight)
    If (lngKeyState And &H8000) Then m_Enemies(0).WorldX = m_Enemies(0).WorldX + 1000
    
    lngKeyState = GetKeyState(vbKeyUp)
    If (lngKeyState And &H8000) Then m_Enemies(0).WorldY = m_Enemies(0).WorldY - 1000
    
    lngKeyState = GetKeyState(vbKeyDown)
    If (lngKeyState And &H8000) Then m_Enemies(0).WorldY = m_Enemies(0).WorldY + 1000
    
    lngKeyState = GetKeyState(vbKeySpace)
    If (lngKeyState And &H8000) Then
        If s_blnKeyDeBounce = False Then
            s_blnKeyDeBounce = True
            g_strGameState = "LevelComplete"
        End If
    Else
        s_blnKeyDeBounce = False
    End If
    
    ' Check for Pause/Resume
    lngKeyState = GetKeyState(vbKeyP)
    If (lngKeyState And &H8000) Then
        If g_strGameState <> "Paused" Then
            s_strPreviousValue = g_strGameState
            g_strGameState = "Paused"
        Else
            g_strGameState = s_strPreviousValue
            frmCanvas.Caption = App.Comments
        End If
    End If
    
    
End Sub

