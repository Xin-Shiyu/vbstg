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

Private MyEnemyPool As New EnemyPool
Private MyBulletPool As New BulletPool
Private EnemyBulletPool As New BulletPool
Private FPSContext As New FPSCalcContext

Private StaticQueue As New TimedActionQueue
Private DynamicQueue As New TimedActionQueue

Private MissTimes As Long
Private Graze As Long

Private GamePaused As Boolean

Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

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
    
    EnemyDamageCheck
    SelfDamageCheck
    
    MyBulletPool.UpdateAll TimeDiff
    MyEnemyPool.UpdateAll TimeDiff
    EnemyBulletPool.UpdateAll TimeDiff
End Sub

Private Sub EnemyDamageCheck()
    Dim ItB As BulletListNode, ItE As EnemyListNode
    Dim Fore As New BulletListNode
    Set Fore.Nx = MyBulletPool.Bullets
    Set ItB = Fore
    Do Until ItB.Nx Is Nothing
        Set ItE = MyEnemyPool.Enemies
        Do Until ItE Is Nothing
            If ItB.Nx.Bullet.GetPosition.DistanceTo(ItE.Enemy.GetPosition) < ItB.Nx.Bullet.GetSize + ItE.Enemy.GetSize Then
                ItE.Enemy.Damage 1
                Set ItB.Nx = ItB.Nx.Nx
                Exit Do
            End If
            Set ItE = ItE.Nx
        Loop
        Set ItB = ItB.Nx
        If ItB Is Nothing Then Exit Do
    Loop
    Set MyBulletPool.Bullets = Fore.Nx
End Sub

Private Sub SelfDamageCheck()
    Dim ItB As BulletListNode
    Dim SelfPos As Point
    Set SelfPos = Self.Pos
    Set ItB = EnemyBulletPool.Bullets
    Do Until ItB Is Nothing
        Dim Distance As Long
        Distance = ItB.Bullet.GetPosition.DistanceTo(SelfPos) - ItB.Bullet.GetSize + 1
        If Distance < 40 Then
            If Distance < 3 Then
                Miss
                Exit Do
            End If
            If Not ItB.Bullet.IsGrazed Then
                ItB.Bullet.SetGrazed
                Graze = Graze + 1
            End If
        End If
        Set ItB = ItB.Nx
    Loop
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
        Context.FillColor = vbWhite
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

Private Sub RenderEnemies()
    Dim It As EnemyListNode
    Dim Pos As Point
    Set It = MyEnemyPool.Enemies
    While Not It Is Nothing
        Set Pos = It.Enemy.GetPosition
        Context.DrawWidth = 1
        Context.FillStyle = vbSolid
        Context.FillColor = vbGreen
        Context.Circle (Pos.X, Pos.Y), It.Enemy.GetSize, vbGreen
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
    Context.Print "Graze:", Graze
    
    RenderEnemies
    RenderMyBullets
    RenderSelf
    RenderEnemyBullets
End Sub

Public Sub Bind(ByVal Frm As Form)
    Set Context = Frm
    Context.Caption = "STG"
    Context.AutoRedraw = True
    Context.BackColor = vbBlack
    Context.Width = Zoom * GameWidth * Screen.TwipsPerPixelX + (Context.Width - Context.ScaleWidth)
    Context.Height = Zoom * GameHeight * Screen.TwipsPerPixelY + (Context.Height - Context.ScaleHeight)
    Context.ScaleWidth = GameWidth
    Context.ScaleHeight = GameHeight
End Sub

Public Sub Run()
    Randomize
    Context.Show
    Dim TimerBegin As Double, TimerEnd As Double
    TimerBegin = 0
    TimerEnd = 1 / 64
    Dim TimeDiff As Double
    StaticQueue.ExecuteLater NewSelfShootingAction(Self, MyBulletPool), 0
    StaticQueue.ExecuteLater NewFPSCalcAction(FPSContext), 1
    Dim i
    For i = 0 To GameWidth Step 50
        Dim MyEnemy As IEnemy
        Set MyEnemy = NewYousei(i, 0 - Rnd * 100)
        MyEnemyPool.Insert MyEnemy
        DynamicQueue.ExecuteLater NewEnemyShootingActionTest(Self, EnemyBulletPool, MyEnemy), 0.5
    Next
    Do While Context.Visible = True
        Dim TimeLeft As Long
        TimeDiff = TimerEnd - TimerBegin
        TimerBegin = Timer
        'If TimeDiff > 1 / 64 Then TimeDiff = 1 / 64
        Update TimeDiff
        Render
        FPSContext.Frames = FPSContext.Frames + 1
        DynamicQueue.Pause
        StaticQueue.Pause
        TimerEnd = Timer
        Do
            DoEvents
        Loop While GamePaused And Context.Visible
        DynamicQueue.Continue
        StaticQueue.Continue
    Loop
    End
End Sub

Private Sub Miss()
    MissTimes = MissTimes + 1
    EnemyBulletPool.Clear
    Render
    Sleep 500
End Sub

Private Sub Bomb()
    
End Sub

Private Sub Context_KeyDown(KeyCode As Integer, Shift As Integer)
    With Self
        Select Case KeyCode
            Case vbKeyX: Bomb
            Case vbKeyZ: .IsShooting = True
            Case vbKeyShift: .SlowMode = True
            Case vbKeyUp: .PressedUp = True
            Case vbKeyDown: .PressedDown = True
            Case vbKeyLeft: .PressedLeft = True
            Case vbKeyRight: .PressedRight = True
            Case vbKeyEscape: End
            Case vbKeySpace: GamePaused = Not GamePaused
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
