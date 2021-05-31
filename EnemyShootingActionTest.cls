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
Public RelatedEnemy As IEnemy
Public Counter As Long

Private Sub IAction_Invoke(ByVal Queue As TimedActionQueue)
    If Not RelatedEnemy.IsAlive Or Counter > 20 Then
        Set RelatedEnemy = Nothing
        Exit Sub
    End If
    
    Dim EPos As Point, SPos As Point, K As Double, Xs As Double, Ys As Double
    Set EPos = RelatedEnemy.GetPosition
    Set SPos = Self.Pos
    Xs = SPos.X - EPos.X
    Ys = SPos.Y - EPos.Y
    K = Sqr(Xs ^ 2 + Ys ^ 2)
    Pool.Insert NewLinear(EPos.X, EPos.Y, Xs / K * 300, Ys / K * 300, 10)
    Counter = Counter + 1
    
    Queue.ExecuteLater Me, 0.2
End Sub

