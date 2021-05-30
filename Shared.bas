Attribute VB_Name = "Shared"
Public Const GameWidth As Double = 576
Public Const GameHeight As Double = 672

Public Function NewLinear(ByVal X As Double, ByVal Y As Double, _
    ByVal Xs As Double, ByVal Ys As Double) As LinearBullet
    Set NewLinear = New LinearBullet
    NewLinear.Position.X = X
    NewLinear.Position.Y = Y
    NewLinear.Speed.X = Xs
    NewLinear.Speed.Y = Ys
End Function

Public Function ShootBullet(ByVal Self As SelfPlane, ByVal Pool As BulletPool)
    With Self.Pos
        Pool.Insert NewLinear(.X, .Y, 0, -2000)
        Pool.Insert NewLinear(.X, .Y, -Self.BulletSpeedX, -2000)
        Pool.Insert NewLinear(.X, .Y, Self.BulletSpeedX, -2000)
    End With
End Function

Public Function NewSelfShootingAction(ByVal Self As SelfPlane, _
    ByVal Pool As BulletPool) As IAction
    Dim Ret As New SelfShootingAction
    Set Ret.Pool = Pool
    Set Ret.Self = Self
    Set NewSelfShootingAction = Ret
End Function

Public Function NewFPSCalcAction(ByVal Context As FPSCalcContext) As IAction
    Dim Ret As New FPSCalcAction
    Set Ret.Context = Context
    Set NewFPSCalcAction = Ret
End Function
