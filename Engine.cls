VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Context As Form
Attribute Context.VB_VarHelpID = -1

Private Self As New SelfPlane

Private ShootingTimer As Double, EnemyTimer As Double
Private MyBulletPool As New BulletPool
Private EnemyBulletPool As New BulletPool
Private FPSContext As New FPSCalcContext

Private StaticQueue As New TimedActionQueue
Private DynamicQueue As New TimedActionQueue

Private MissTimes As Long

Private Static Function TimeUp(ByRef TimerObject As Double, _
    ByVal Standard As Double) As Boolean
    If TimerObject > Standard Then
        TimerObject = TimerObject - Standard
        TimeUp = True
        Exit Function
    End If
    TimeUp = False
End Function

Private Sub Update(ByVal TimeDiff As Double)
    Self.Update TimeDiff
    
    DynamicQueue.Flush
    StaticQueue.Flush
    
    MyBulletPool.UpdateAll TimeDiff
    EnemyBulletPool.UpdateAll TimeDiff
End Sub

Private Sub RenderMyBullets()
    Dim It As BulletListNode
    Dim Pos As Point
    Set It = MyBulletPool.Bullets
    While Not (It Is Nothing)
        Set Pos = It.Bullet.GetPosition
        Context.DrawWidth = 2
        Context.FillStyle = vbCross
        Context.FillColor = vbRed
        Context.Circle (Pos.X, Pos.Y), It.Bullet.GetSize, vbRed
        Set It = It.Nx
    Wend
End Sub

Private Sub RenderSelf()
    With Self
        Context.DrawWidth = 3
        Context.FillStyle = vbSolid
        Dim BulletDistance As Double
        BulletDistance = EnemyBulletPool.MinDistance(.Pos.X, .Pos.Y)
        If BulletDistance > 28.28 Then
            Context.FillColor = vbWhite
        ElseIf BulletDistance > 5 Then
            Context.FillColor = vbYellow
        Else
            Context.FillColor = vbBlack
            MissTimes = MissTimes + 1
            EnemyBulletPool.Clear
        End If
        Context.Circle (.Pos.X, .Pos.Y), 20, vbRed
        Context.DrawWidth = 1
        If Self.SlowMode Then
            Context.Circle (.Pos.X, .Pos.Y), 5, vbBlack
        End If
    End With
End Sub

Private Sub RenderEnemyBullets()
    Dim It As BulletListNode
    Dim Pos As Point
    Set It = EnemyBulletPool.Bullets
    While Not It Is Nothing
        Set Pos = It.Bullet.GetPosition
        Context.DrawWidth = 1
        Context.FillStyle = vbSolid
        Context.FillColor = vbWhite
        Context.Circle (Pos.X, Pos.Y), It.Bullet.GetSize, vbBlue
        Set It = It.Nx
    Wend
End Sub

Private Sub Render()
    Context.ForeColor = vbWhite
    Context.Cls
    
    Context.Print "FPS:", FPSContext.FPS
    Context.Print "My Bullet Count:", MyBulletPool.BulletCount
    Context.Print "Enemy Bullet Count:", EnemyBulletPool.BulletCount
    Context.Print "Miss Times:", MissTimes
    
    RenderMyBullets
    RenderSelf
    RenderEnemyBullets
End Sub

Public Sub Bind(ByVal Frm As Form)
    Set Context = Frm
    Context.Caption = "STG"
    Context.AutoRedraw = True
    Context.BackColor = vbBlack
    Context.Width = GameWidth * Screen.TwipsPerPixelX + (Context.Width - Context.ScaleWidth)
    Context.Height = GameHeight * Screen.TwipsPerPixelY + (Context.Height - Context.ScaleHeight)
    Context.ScaleMode = vbPixels
End Sub

Public Sub Run()
    Randomize
    Context.Show
    Dim LastTimer As Double
    LastTimer = Timer
    StaticQueue.ExecuteLater NewSelfShootingAction(Self, MyBulletPool), 0
    StaticQueue.ExecuteLater NewFPSCalcAction(FPSContext), 1
    DynamicQueue.ExecuteLater NewEnemyShootingActionTest(Self, EnemyBulletPool), 0.5
    Do While Context.Visible = True
        Dim TimeDiff As Double
        TimeDiff = Timer - LastTimer
        Update TimeDiff
        LastTimer = Timer
        Render
        FPSContext.Frames = FPSContext.Frames + 1
        DoEvents
    Loop
    End
End Sub

Private Sub Context_KeyDown(KeyCode As Integer, Shift As Integer)
    With Self
        Select Case KeyCode
            Case vbKeyZ: .IsShooting = True
            Case vbKeyShift: .SlowMode = True
            Case vbKeyUp: .PressedUp = True
            Case vbKeyDown: .PressedDown = True
            Case vbKeyLeft: .PressedLeft = True
            Case vbKeyRight: .PressedRight = True
            Case vbKeyEscape: End
        End Select
    End With
End Sub

Private Sub Context_KeyUp(KeyCode As Integer, Shift As Integer)
    With Self
        Select Case KeyCode
            Case vbKeyZ: .IsShooting = False
            Case vbKeyShift: .SlowMode = False
            Case vbKeyUp: .PressedUp = False
            Case vbKeyDown: .PressedDown = False
            Case vbKeyLeft: .PressedLeft = False
            Case vbKeyRight: .PressedRight = False
        End Select
    End With
End Sub
