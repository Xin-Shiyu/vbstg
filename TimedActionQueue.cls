VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimedActionQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Front As ActionQueueNode
Private Back As ActionQueueNode
Private TimeStopped As Double
Private LastStop As Double

Public Function Timing() As Single
    Timing = Timer - TimeStopped
End Function

Public Sub ExecuteLater(Action As IAction, Delay As Single)
    Dim ActualTime As Single
    ActualTime = Timer - TimeStopped + Delay
    Dim Node As New ActionQueueNode
    Set Node.Action = Action
    Node.When = ActualTime
    If Back Is Nothing Or Front Is Nothing Then
        Debug.Assert Front Is Back
        Set Front = Node
        Set Back = Node
    Else
        Dim It As ActionQueueNode
        Set It = Back
        Debug.Assert Not Back Is Nothing
        Do Until It.When < ActualTime
            If It Is Front Then
                Set Front.Pv = Node
                Set Node.Nx = Front
                Set Front = Node
                Exit Sub
            End If
            Set It = It.Pv
        Loop
        Set Node.Nx = It.Nx
        Set Node.Pv = It
        Set It.Nx = Node
        If It Is Back Then Set Back = Node
        Debug.Assert Not Front.Nx Is Nothing
    End If
End Sub

Public Sub Flush()
    Dim PendingAction As New Collection
    Do Until Front Is Nothing
        If Front.When > Timer - TimeStopped Then
            Exit Do
        End If
        PendingAction.Add Front.Action
        Set Front = Front.Nx
        If Front Is Nothing Then
            Set Back = Nothing
        Else
            Set Front.Pv = Nothing
        End If
    Loop
    Dim Action
    For Each Action In PendingAction
        Action.Invoke Me
    Next
End Sub

Public Sub Pause()
    LastStop = Timer
End Sub

Public Sub Continue()
    TimeStopped = TimeStopped + Timer - LastStop
End Sub
