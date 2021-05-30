VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnemyShootingActionTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IAction

Public Pool As BulletPool
Public Self As SelfPlane

Private Sub IAction_Invoke(ByVal Queue As TimedActionQueue)
    Dim I As Double, Y As Double
    Dim Target As Integer
    Target = Rnd * 60 - 30
    For I = -200 To 200 Step 10
        Y = Sqr(200 ^ 2 - I ^ 2)
        If Abs(I - Target) >= 30 Then
            Pool.Insert NewLinear(GameWidth / 2, 0, I, Y, 8)
        End If
    Next
    Queue.ExecuteLater NewEnemyShootingActionTest(Self, Pool), 0.1
End Sub
