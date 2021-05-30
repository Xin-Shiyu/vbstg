VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelfPlane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Pos As New Point
Public PressedUp As Boolean, PressedDown As Boolean
Public PressedLeft As Boolean, PressedRight As Boolean
Public SlowMode As Boolean, IsShooting As Boolean
Public BulletSpeedX As Double

Private Const SlowSpeed = 100, FastSpeed = 350

Public Sub Update(ByVal TimeDiff As Double)
    Dim MoveDiff As Double
    If SlowMode Then
        MoveDiff = SlowSpeed * TimeDiff
    Else
        MoveDiff = FastSpeed * TimeDiff
    End If
    With Pos
        If PressedUp Then .Y = .Y - MoveDiff
        If PressedDown Then .Y = .Y + MoveDiff
        If PressedLeft Then .X = .X - MoveDiff
        If PressedRight Then .X = .X + MoveDiff
        If .Y > GameHeight Then .Y = GameHeight
        If .X > GameWidth Then .X = GameWidth
        If .Y < 0 Then .Y = 0
        If .X < 0 Then .X = 0
    End With
End Sub
