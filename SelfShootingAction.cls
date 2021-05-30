VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelfShootingAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IAction

Public Pool As BulletPool
Public Self As SelfPlane

Private Sub IAction_Invoke(ByVal Queue As TimedActionQueue)
    If Self.IsShooting Then ShootBullet Self, Pool
    If Self.SlowMode Then
        Self.BulletSpeedX = Self.BulletSpeedX + (10 - Self.BulletSpeedX) / 2
    Else
        Self.BulletSpeedX = Self.BulletSpeedX + (1000 - Self.BulletSpeedX) / 2
    End If
    Queue.ExecuteLater NewSelfShootingAction(Self, Pool), 0.01
End Sub
