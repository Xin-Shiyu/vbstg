VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Yousei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IEnemy

Public Pos As New Point
Public Life As Long

Private Sub Class_Initialize()
    Life = 25
End Sub

Private Sub IEnemy_Damage(ByVal Extent As Long)
    If Life <= 0 Then Exit Sub
    Life = Life - Extent
End Sub

Private Function IEnemy_GetPosition() As Point
    Set IEnemy_GetPosition = Pos
End Function

Private Function IEnemy_GetSize() As Double
    IEnemy_GetSize = 20
End Function

Private Function IEnemy_IsAlive() As Boolean
    IEnemy_IsAlive = Life > 0 And Pos.Y < GameHeight + 20
End Function

Private Sub IEnemy_Update(ByVal TimeDiff As Double)
    Pos.Y = Pos.Y + TimeDiff * 30
End Sub
