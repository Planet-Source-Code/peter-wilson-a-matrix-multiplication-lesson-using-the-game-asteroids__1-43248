Attribute VB_Name = "mGameLoop"
Option Explicit

Public g_strGameState As String
Private m_lngApplicationState As Long
Private m_lngDifficulty As Long
Private m_intLevel As Integer


Private m_PlayerShip As mdr2DObject
Public m_Enemies() As mdr2DObject
Private m_GameObjects() As mdr2DObject

Private m_MaxAsteroids As Integer
Private m_Asteroids() As mdr2DObject

' Particles are objects that have a limited life span.
Private m_MaxParticles As Integer
Private m_Particles() As mdr2DObject

' World Window Limits ie. This is the game's world coordinates (which could be very large)
Private m_Xmin As Single
Private m_Xmax As Single
Private m_Ymin As Single
Private m_Ymax As Single

' ViewPort Limits ie. Usually the limits of a VB form, or picturebox (which could be very small)
Private m_Umin As Single
Private m_Umax As Single
Private m_Vmin As Single
Private m_Vmax As Single

' Module Level Matrices (that don't change much)
Private m_matScale As mdrMATRIX3x3
Private m_matViewMapping As mdrMATRIX3x3



Public Function Create_Particles(NumberOfParticles As Integer, MinSize As Integer, MaxSize As Integer, WorldX As Single, WorldY As Single, VectorX As Single, VectorY As Single, Red As Integer, Green As Integer, Blue As Integer, LifeTime As Single) As Integer

    ' "Attempts to create" the specified number of particles,
    ' and returns the number of particles "actually created".
    Create_Particles = 0
    
    Dim intN As Integer
    Dim sngRadius As Single
    
    intN = 0
    Do
        If m_Particles(intN).Enabled = False Then
            ' This particle is no longer used, so we can use this one --> m_Particles(intN)
            
            If NumberOfParticles > 0 Then
                
                ' Create a random sized asteroid within the min/max parameters specified.
                sngRadius = GetRNDNumberBetween(MinSize, MaxSize)
                m_Particles(intN) = CreateRandomShapeAsteroid(sngRadius)
                
                ' Fill-in some properties.
                With m_Particles(intN)
                    .Enabled = True
                    .Visible = True
                    .Caption = "Particle"
                    .ParticleLifeRemaining = LifeTime
                    
                    .WorldX = WorldX
                    .WorldY = WorldY
                    
                    ' Initial Vector
                    If VectorX = 0 Then
                        .Vector.x = GetRNDNumberBetween(-700, 700)
                    Else
                        .Vector.x = VectorX
                    End If
                    
                    If VectorY = 0 Then
                        .Vector.y = GetRNDNumberBetween(-700, 700)
                    Else
                        .Vector.y = VectorY
                    End If
                    .Vector.w = 1
                    
                    .Mass = sngRadius
                    
                    .SpinVector = GetRNDNumberBetween(-4, 4)
                    .RotationAboutZ = 0
                    
                    .Red = Red: .Green = Green: .Blue = Blue
                    
                End With
                
                NumberOfParticles = NumberOfParticles - 1
            End If
        End If
        
        intN = intN + 1
    Loop Until (intN = m_MaxParticles) Or (NumberOfParticles = 0)
        
End Function
Public Function Create_Asteroids(ByVal Qty As Integer, MinSize As Integer, MaxSize As Integer, WorldX As Single, WorldY As Single, Red As Integer, Green As Integer, Blue As Integer, LifeTime As Single) As Integer

    ' "Attempts to create" the specified number of Asteroids,
    ' and returns the number of Asteroids "actually created".
    Create_Asteroids = 0
    
    Dim intN As Integer
    Dim sngRadius As Single
    
    intN = 0
    Do
        If m_Asteroids(intN).Enabled = False Then
            ' This Asteroid is no longer used, so we can use this one --> m_Asteroids(intN)
            
            If Qty > 0 Then
                
                ' Create a random sized asteroid within the min/max parameters specified.
                sngRadius = GetRNDNumberBetween(MinSize, MaxSize)
                m_Asteroids(intN) = CreateRandomShapeAsteroid(sngRadius)
                
                ' Fill-in some properties.
                With m_Asteroids(intN)
                    .Enabled = True
                    .Visible = True
                    .Caption = "Asteroid"
                    .ParticleLifeRemaining = LifeTime
                    
                    .WorldX = WorldX
                    .WorldY = WorldY
                    
                    ' Initial Vector
                    .Vector.x = GetRNDNumberBetween(-700, 700)
                    .Vector.y = GetRNDNumberBetween(-700, 700)
                    .Vector.w = 1
                    
                    .SpinVector = GetRNDNumberBetween(-4, 4)
                    .RotationAboutZ = 0
                    
                    .Red = Red: .Green = Green: .Blue = Blue
                    
                End With
                
                Qty = Qty - 1
            End If
        End If
        
        intN = intN + 1
    Loop Until (intN = m_MaxParticles) Or (Qty = 0)
        
End Function

Private Sub GameIsPaused()

    Static s_lngCounter As Long
    Static s_blnFlipFlop As Boolean
    
    s_lngCounter = s_lngCounter + 1
    If s_lngCounter = (2 ^ 31) - 1 Then s_lngCounter = 0
    
    If (s_lngCounter Mod 10) = 0 Then s_blnFlipFlop = Not s_blnFlipFlop
    
    If s_blnFlipFlop = True Then
        frmCanvas.Caption = "* * * P A U S E D * * *"
    Else
        frmCanvas.Caption = "* * *   G A M E   * * *"
    End If
    
End Sub

Public Sub Initilize_Game_Data()

    ReDim g_Asteroids(0)
    g_Asteroids(0) = Create_TestTriangle
    
    ReDim g_SpaceShips(0)
    g_SpaceShips(0) = Create_PlayerSpaceShip
    g_SpaceShips(0).WorldX = 5
    g_SpaceShips(0).WorldY = 0
    
End Sub

Public Sub Main()

    ' ==========================================================================
    ' This routine get's called by a Timer Event regardless of what's happening.
    ' (Although you can have multiple Timer Controls, it tends to make programs
    '  disorganised and less predictable. By using only a single Timer control,
    '  I have very strict control over what occurs and when. This routine is
    '  actually a mini-"state machine"... well actually most computer programs
    '  are, but I digress... look them up, learn them, they are cool.)
    '  Also See: www.midar.com.au/vblessons/
    ' ==========================================================================
    
    Call ProcessKeyboardInput
    
    Select Case g_strGameState
        Case ""
            Randomize
            m_intLevel = 0
            m_lngDifficulty = 0
            
            m_MaxParticles = 60
            ReDim m_Particles(m_MaxParticles)
            
            frmCanvas.Show
            g_strGameState = "LevelComplete"
            
        Case "PlayingLevel"
            Call PlayGame
            
        Case "LevelComplete"
            m_intLevel = m_intLevel + 1
            Call LoadLevel(m_intLevel, m_lngDifficulty)
            g_strGameState = "PlayingLevel"
            
        Case "Paused"
            Call GameIsPaused
            
    End Select
    
End Sub
Private Sub LoadLevel(Level As Integer, Difficulty As Long)

    Dim intN As Integer
    Dim sngRadius As Single
        
    ' ================
    ' Create Asteroids
    ' ================
    ' One Large Asteroid can be split into two medium asteroids,
    ' and then each of these medium ones, can be split again into smaller ones.
    m_MaxAsteroids = Level * 4
    If m_MaxAsteroids <> 0 Then
        ReDim m_Asteroids(m_MaxAsteroids - 1)
        Call Create_Asteroids(Level, 1000, 1000, 0, 0, 0, 192, 192, 0)
    End If
    
    
    ' =====================================================================
    ' Create Enemies
    ' This should be space ships, but I've just made them asteroids for now
    ' =====================================================================
    ReDim m_Enemies(Int(Level / 2))
    For intN = 0 To Int(Level / 2)
        sngRadius = GetRNDNumberBetween(200, 800)
        m_Enemies(intN) = CreateRandomShapeAsteroid(sngRadius) ' Create_PlayerSpaceShip
        With m_Enemies(intN)
            .Caption = "Enemy" & intN
            .Enabled = True
            .WorldX = GetRNDNumberBetween(-20000, 20000)
            .WorldY = GetRNDNumberBetween(-20000, 20000)
            .Vector.x = 0
            .Vector.y = 0
            .SpinVector = GetRNDNumberBetween(-5, 5)
            .Mass = sngRadius
            .Red = 255: .Green = 0: .Blue = 0
        End With
    Next intN
    
    
    
End Sub

Public Sub PlayGame()

    Static s_lngCounter As Long
    s_lngCounter = s_lngCounter + 1
    If s_lngCounter = (2 ^ 31) - 1 Then s_lngCounter = 0
    
    ' Try change these values.
    m_matScale = MatrixScaling(2, 2)
    
    Dim intParticlesCreated As Integer
    
    Call Calculate_Asteroids
    
    intParticlesCreated = Calculate_Particles
    
    Call Calculate_Enemies_Part1
    
    Call Refresh_GameScreen
    
    Call Calculate_Enemies_Part2
    Call Calculate_Enemies_Part2b
    
    Select Case s_lngCounter
        Case 100
            MsgBox "Press the Space Bar to change levels", vbInformation
            
        Case 300
            MsgBox "Use the arrow keys to move one of the little red things around", vbInformation
            
    End Select
    
End Sub


Public Sub Calculate_Asteroids()

    Dim intN As Integer
    Dim intJ As Integer
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
    
    On Error GoTo errTrap
    
    
    For intN = LBound(m_Asteroids) To UBound(m_Asteroids)
        With m_Asteroids(intN)
            If .Enabled = True Then
            
                ' Translate
                ' =========
                .WorldX = .WorldX + .Vector.x
                .WorldY = .WorldY + .Vector.y
                
                If .WorldX > m_Xmax Then .WorldX = .WorldX - (m_Xmax - m_Xmin)
                If .WorldX < m_Xmin Then .WorldX = .WorldX + (m_Xmax - m_Xmin)
                If .WorldY > m_Ymax Then .WorldY = .WorldY - (m_Ymax - m_Ymin)
                If .WorldY < m_Ymin Then .WorldY = .WorldY + (m_Ymax - m_Ymin)
                
                matTranslate = MatrixTranslation(.WorldX, .WorldY)
                
                .RotationAboutZ = .RotationAboutZ + .SpinVector
                matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))
    
                
                ' Multiply matrices in the correct order.
                matResult = MatrixIdentity
                matResult = MatrixMultiply(matResult, m_matScale)
                matResult = MatrixMultiply(matResult, matRotationAboutZ)
                matResult = MatrixMultiply(matResult, matTranslate)
                matResult = MatrixMultiply(matResult, m_matViewMapping)
                
                For intJ = LBound(.Vertex) To UBound(.Vertex)
                    .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
                Next intJ
                
                ' Conditionally Compiled (see Project Properties)
                #If gcShowVectors = -1 Then
                    ' Transform the Direction/Speed Vector to screen space
                    ' Do this step, only if you wish to display this vector on the screen.
                    ' Displaying the vector on screen, is only useful for debugging/instructional purposes.
                    ' Remember, DO NOT rotate the Direction/Speed vector (try it, and see what happens)
                    matResult = MatrixIdentity
                    matResult = MatrixMultiply(matResult, m_matScale)
                    ''matResult = MatrixMultiply(matResult, matRotationAboutZ)
                    matResult = MatrixMultiply(matResult, matTranslate)
                    matResult = MatrixMultiply(matResult, m_matViewMapping)
                    .TVector = MatrixMultiplyVector(matResult, .Vector)
                #End If
                
            End If ' Is Enabled?
        End With
    Next intN

    Exit Sub
errTrap:

End Sub
Public Function Calculate_Particles() As Long

    ' Processes all Particles, then returns the number of active particles.
    
    Dim intN As Integer
    Dim intJ As Integer
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
    
    Calculate_Particles = 0
    
    For intN = LBound(m_Particles) To UBound(m_Particles)
        With m_Particles(intN)
            If .Enabled = True Then
                Calculate_Particles = Calculate_Particles + 1
                
                ' Translate
                ' =========
                .WorldX = .WorldX + .Vector.x
                .WorldY = .WorldY + .Vector.y
                
                If .WorldX > m_Xmax Then .WorldX = .WorldX - (m_Xmax - m_Xmin)
                If .WorldX < m_Xmin Then .WorldX = .WorldX + (m_Xmax - m_Xmin)
                If .WorldY > m_Ymax Then .WorldY = .WorldY - (m_Ymax - m_Ymin)
                If .WorldY < m_Ymin Then .WorldY = .WorldY + (m_Ymax - m_Ymin)
                
                matTranslate = MatrixTranslation(.WorldX, .WorldY)
                
                .RotationAboutZ = .RotationAboutZ + .SpinVector
                matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))
    
                
                ' Multiply matrices in the correct order.
                matResult = MatrixIdentity
                matResult = MatrixMultiply(matResult, m_matScale)
                matResult = MatrixMultiply(matResult, matRotationAboutZ)
                matResult = MatrixMultiply(matResult, matTranslate)
                matResult = MatrixMultiply(matResult, m_matViewMapping)
                
                For intJ = LBound(.Vertex) To UBound(.Vertex)
                    .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
                Next intJ
                
                ' Transform the Direction/Speed Vector to screen space
                ' Do this step, only if you wish to display this vector on the screen.
                ' Displaying the vector on screen, is only useful for debugging/instructional purposes.
                ' Remember, DO NOT rotate the Direction/Speed vector (try it, and see what happens)
                matResult = MatrixIdentity
                matResult = MatrixMultiply(matResult, m_matScale)
                ''matResult = MatrixMultiply(matResult, matRotationAboutZ)
                matResult = MatrixMultiply(matResult, matTranslate)
                matResult = MatrixMultiply(matResult, m_matViewMapping)
                .TVector = MatrixMultiplyVector(matResult, .Vector)
                
                
                
                ' Reduce Particle life
                .ParticleLifeRemaining = .ParticleLifeRemaining - 1
                If .ParticleLifeRemaining < 1 Then .Enabled = False
                
                
                ' Fade to Dull Red, then to black.
                .Red = .Red - 6
                .Green = .Green - 16
                .Blue = .Blue - 16
                
                
            End If ' Is Enabled?
            
        End With
    Next intN

End Function

Public Sub Calculate_Enemies_Part1()

    Dim intN As Integer
    Dim intJ As Integer
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
    
    For intN = LBound(m_Enemies) To UBound(m_Enemies)
        With m_Enemies(intN)
                        
            ' Translate
            ' =========
            .WorldX = .WorldX + .Vector.x
            .WorldY = .WorldY + .Vector.y
            
            If .WorldX > m_Xmax Then .WorldX = .WorldX - (m_Xmax - m_Xmin)
            If .WorldX < m_Xmin Then .WorldX = .WorldX + (m_Xmax - m_Xmin)
            If .WorldY > m_Ymax Then .WorldY = .WorldY - (m_Ymax - m_Ymin)
            If .WorldY < m_Ymin Then .WorldY = .WorldY + (m_Ymax - m_Ymin)
            
            matTranslate = MatrixTranslation(.WorldX, .WorldY)
            
            .RotationAboutZ = .RotationAboutZ + .SpinVector
            matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))

            
            ' Multiply matrices in the correct order.
            matResult = MatrixIdentity
            matResult = MatrixMultiply(matResult, m_matScale)
            matResult = MatrixMultiply(matResult, matRotationAboutZ)
            matResult = MatrixMultiply(matResult, matTranslate)
            matResult = MatrixMultiply(matResult, m_matViewMapping)
            
            For intJ = LBound(.Vertex) To UBound(.Vertex)
                .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
            Next intJ
            
        End With
    Next intN

End Sub

Public Sub Calculate_Enemies_Part2()

    Dim intEnemy As Integer
    Dim intAsteroid As Integer
    Dim tempV As mdrVector3
    Dim tempV3 As mdrVector3
    Dim VDisplay As mdrVector3
    Dim tempV2 As mdrVector3
    Dim sngDistance As Single
    Dim sngMultiplier As Single
    
    On Error GoTo errTrap
    
    For intEnemy = LBound(m_Enemies) To UBound(m_Enemies)
        With m_Enemies(intEnemy)
            If .Enabled = True Then
                For intAsteroid = LBound(m_Asteroids) To UBound(m_Asteroids)
                    If m_Asteroids(intAsteroid).Enabled = True Then
                    
                    tempV.x = .WorldX - m_Asteroids(intAsteroid).WorldX
                    tempV.y = .WorldY - m_Asteroids(intAsteroid).WorldY
                    tempV.w = 1
                    
                    sngDistance = Vec3Length(tempV)
                    If (sngDistance < 55000) Then
                        tempV = Vec3Normalize(tempV)
                        
                        #If gcShowVectors = -1 Then
                            VDisplay = tempV
                            VDisplay = Vec3MultiplyByScalar(VDisplay, 3000)
                            VDisplay.x = .WorldX + VDisplay.x
                            VDisplay.y = .WorldY + VDisplay.y
                        #End If
    
                        sngMultiplier = (1 - (sngDistance / m_Xmax)) * 1000 ' <<< Change this * 1000 bit!!!
                        tempV = Vec3MultiplyByScalar(tempV, sngMultiplier)
                        
                        .WorldX = .WorldX + tempV.x
                        .WorldY = .WorldY + tempV.y
                        
                        
                        #If gcShowVectors = -1 Then
                            VDisplay = MatrixMultiplyVector(m_matViewMapping, VDisplay)
                            frmCanvas.ForeColor = RGB(255, 0, 255)
                            frmCanvas.DrawWidth = 2
                            frmCanvas.PSet (VDisplay.x, VDisplay.y)
    '                        frmCanvas.Font = "Small Fonts"
    '                        frmCanvas.FontSize = 7
    '                        frmCanvas.Print Format(sngMultiplier, "0.00000")
                        #End If
                    End If ' Is Asteroid close to us?
                    
                    
                    ' =====================================================================
                    ' This is a VERY fun place to change paramters!
                    ' Minuses / Pluses. world coordininates, local, it just doesn't matter!
                    ' =====================================================================
                    If (sngDistance < 9000) Then
                        tempV3 = Vec3Normalize(tempV)
                        tempV3 = Vec3MultiplyByScalar(tempV3, 1900)
                        Call Create_Particles(1, 100, 100, .WorldX, .WorldY, tempV3.x, tempV3.y, 255, 255, 0, 10)
'                        Call Create_Particles(1, 100, 100, .WorldX, .WorldY, 0, 0, 255, 255, 0, 10)
                        Call Create_Particles(1, CInt(.Mass), CInt(.Mass), .WorldX, .WorldY, 0, 0, 192, 127, 127, 15)
                    End If
                    
                    
                    End If ' Is Asteroid Enabled?
                Next intAsteroid
            End If ' Is Enemy Enabled?
        End With
    Next intEnemy
    
    Exit Sub
errTrap:
    
End Sub
Public Sub Calculate_Enemies_Part2b()

    Dim intEnemy As Integer
    Dim intOtherEnemy As Integer
    Dim tempV As mdrVector3
    Dim VDisplay As mdrVector3
    Dim tempV2 As mdrVector3
    Dim sngDistance As Single
    
    Dim sngMultiplier As Single
        
    For intEnemy = LBound(m_Enemies) To UBound(m_Enemies)
        With m_Enemies(intEnemy)
            If .Enabled = True Then
                For intOtherEnemy = LBound(m_Enemies) To UBound(m_Enemies)
                    If (m_Enemies(intOtherEnemy).Enabled = True) And (intOtherEnemy <> intEnemy) Then
                    
                            tempV.x = .WorldX - m_Enemies(intOtherEnemy).WorldX
                            tempV.y = .WorldY - m_Enemies(intOtherEnemy).WorldY
                            tempV.w = 1
                            
                            sngDistance = Vec3Length(tempV)
                            If (sngDistance > 1) Then
                                tempV = Vec3Normalize(tempV)
                                
                                #If gcShowVectors = -1 Then
                                    VDisplay = tempV
                                    VDisplay = Vec3MultiplyByScalar(VDisplay, 1500)
                                    VDisplay.x = .WorldX + VDisplay.x
                                    VDisplay.y = .WorldY + VDisplay.y
                                #End If
            
                                sngMultiplier = (1 - (sngDistance / m_Xmax)) * 400 ' <<< Change this * 400 bit!!!
            '                    If sngMultiplier < 0 Then sngMultiplier = -1
                                tempV = Vec3MultiplyByScalar(tempV, sngMultiplier)
                                
                                .WorldX = .WorldX + tempV.x
                                .WorldY = .WorldY + tempV.y
                                
                                
                                #If gcShowVectors = -1 Then
                                    VDisplay = MatrixMultiplyVector(m_matViewMapping, VDisplay)
                                    frmCanvas.ForeColor = RGB(0, 255, 127)
                                    frmCanvas.DrawWidth = 2
                                    frmCanvas.PSet (VDisplay.x, VDisplay.y)
            '                        frmCanvas.Font = "Small Fonts"
            '                        frmCanvas.FontSize = 7
            '                        frmCanvas.Print Format(sngMultiplier, "0.00000")
                                #End If
                            End If ' Is Asteroid close to us?
                                        
                    End If ' Is Other Enemy Enabled?
                Next intOtherEnemy
            End If ' Is Enemy Enabled?
        End With
    Next intEnemy
    

End Sub
Public Sub Draw_Faces(CurrentObject() As mdr2DObject)

    Dim intN As Integer
    Dim intFaceIndex As Integer
    Dim intK As Integer
    Dim intVertexIndex As Integer
    Dim xPos As Single
    Dim yPos As Single
    Dim matTranslate As mdrMATRIX3x3
    Dim tempXPos As Single
    Dim tempYPos As Single
    
    On Error GoTo errTrap
    
    frmCanvas.DrawMode = vbCopyPen
    frmCanvas.DrawWidth = 1
    
    For intN = LBound(CurrentObject) To UBound(CurrentObject)
        With CurrentObject(intN)
            
            If .Enabled = True Then
            
                ' Clamp values to safe levels
                If .Red < 0 Then .Red = 0
                If .Green < 0 Then .Green = 0
                If .Blue < 0 Then .Blue = 0
                If .Red > 255 Then .Red = 255
                If .Green > 255 Then .Green = 255
                If .Blue > 255 Then .Blue = 255
                
                ' Set colour of Object
                frmCanvas.ForeColor = RGB(.Red, .Green, .Blue)
                
                For intFaceIndex = LBound(.Face) To UBound(.Face)
                    
                    For intK = LBound(.Face(intFaceIndex)) To UBound(.Face(intFaceIndex))
                    
                        intVertexIndex = .Face(intFaceIndex)(intK)
                        xPos = .TVertex(intVertexIndex).x
                        yPos = .TVertex(intVertexIndex).y
                        
                        If LBound(.Face(intFaceIndex)) = UBound(.Face(intFaceIndex)) Then
                            If .Caption = "Asteroid" Then
                                ' This is a face having only a single dot.
                                ' This area is useful for debuggin purposes.
    ''                            frmCanvas.Font = "Small Fonts"
    ''                            frmCanvas.FontSize = 7
    ''                            frmCanvas.CurrentX = xPos
    ''                            frmCanvas.CurrentY = yPos
    '''                            frmCanvas.Print Int(.WorldX) & ", " & Int(.WorldY)
    '''                            frmCanvas.Print .Mass
                                ' Conditionally Compiled (see Project Properties)
                                #If gcShowVectors = -1 Then
                                    frmCanvas.DrawWidth = 3
                                    frmCanvas.ForeColor = RGB(255, 255, 0)
                                    frmCanvas.PSet (xPos, yPos)
                                    frmCanvas.DrawWidth = 1
                                    frmCanvas.ForeColor = RGB(.Red, .Green, .Blue)
                                    frmCanvas.Line (xPos, yPos)-(.TVector.x, .TVector.y)
                                #End If
                            End If
                        Else
                        
                            ' Normal Face; move to first point, then draw to the others.
                            ' ==========================================================
                            If intK = LBound(.Face(intFaceIndex)) Then
                                ' Move to first point
                                frmCanvas.Line (xPos, yPos)-(xPos, yPos)
                                
                            Else
                                ' Draw to point
                                frmCanvas.Line -(xPos, yPos)
                            End If
                            
                        End If
                        
                    Next intK
                Next intFaceIndex
                
                
            End If ' Is Enabled?
        End With
    Next intN

    Exit Sub
errTrap:

End Sub
Public Sub Init_ViewMapping()

    ' Set the size of the World's window.
    m_Xmin = -32768
    m_Xmax = 32767
    m_Ymin = -32768
    m_Ymax = 32767
    
    ' Set the size of the ViewPort's windows.
    m_Umin = 0
    m_Umax = frmCanvas.Width
    m_Vmin = 0
    m_Vmax = frmCanvas.Height
    
    m_matViewMapping = MatrixViewMapping(m_Xmin, m_Xmax, m_Ymin, m_Ymax, m_Umin, m_Umax, m_Vmin, m_Vmax)

End Sub


Private Sub Refresh_GameScreen()

    frmCanvas.Cls
    Call DrawCrossHairs
    Call Draw_Faces(m_Asteroids)
    Call Draw_Faces(m_Particles)
    Call Draw_Faces(m_Enemies)
    
End Sub

