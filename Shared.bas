Attribute VB_Name = "Shared"
Public Const GameWidth As Double = 576
Public Const GameHeight As Double = 672
Public Const SelfBulletSpeed = 1344
Public Const SelfBulletSpeedXMax = 950
Public Const SelfBulletSpeedXMin = 10
Public Const Zoom = 1

Public Function NewLinear(ByVal X As Double, ByVal Y As Double, _
    ByVal Xs As Double, ByVal Ys As Double, Optional Size As Double = 5) As LinearBullet
    Set NewLinear = New LinearBullet
    NewLinear.Position.X = X
    NewLinear.Position.Y = Y
    NewLinear.Speed.X = Xs
    NewLinear.Speed.Y = Ys
    NewLinear.Size = Size
End Function

Public Function ShootBullet(ByVal Self As SelfPlane, ByVal Pool As BulletPool)
    With Self.Pos
        Dim SlantY As Double
        SlantY = Sqr(1 - (Self.BulletSpeedX / SelfBulletSpeed) ^ 2) * SelfBulletSpeed
        Pool.Insert NewLinear(.X, .Y, 0, -SelfBulletSpeed)
        Pool.Insert NewLinear(.X, .Y, -Self.BulletSpeedX, -SlantY)
        Pool.Insert NewLinear(.X, .Y, Self.BulletSpeedX, -SlantY)
    End With
End Function

Public Function NewSelfShootingAction(ByVal Self As SelfPlane, _
    ByVal Pool As BulletPool) As IAction
    Dim Ret As New SelfShootingAction
    Set Ret.Pool = Pool
    Set Ret.Self = Self
    Set NewSelfShootingAction = Ret
End Function

Public Function NewEnemyShootingActionTest(ByVal Self As SelfPlane, _
    ByVal Pool As BulletPool, ByVal RelatedEnemy As IEnemy) As IAction
    Dim Ret As New EnemyShootingActionTest
    Set Ret.Pool = Pool
    Set Ret.Self = Self
    Set Ret.RelatedEnemy = RelatedEnemy
    Set NewEnemyShootingActionTest = Ret
End Function

Public Function NewFPSCalcAction(ByVal Context As FPSCalcContext) As IAction
    Dim Ret As New FPSCalcAction
    Set Ret.Context = Context
    Set NewFPSCalcAction = Ret
End Function

Public Function NewYousei(ByVal X As Double, ByVal Y As Double) As Yousei
    Dim Ret As New Yousei
    Ret.Pos.X = X
    Ret.Pos.Y = Y
    Set NewYousei = Ret
End Function
