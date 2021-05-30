VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StaticEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IEnemy

Public Pos As New Point

Private Sub Class_Initialize()
    Pos.X = GameWidth / 2
    Pos.Y = GameWidth / 2
End Sub

Private Sub IEnemy_Damage(Extent As Long)
    
End Sub

Private Function IEnemy_GetPosition() As Point
    IEnemy_GetPosition = Pos
End Function

Private Function IEnemy_GetSize() As Double
    IEnemy_GetSize = 20
End Function

Private Function IEnemy_IsAlive() As Boolean
    IEnemy_IsAlive = True
End Function
