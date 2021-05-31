VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnemyPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enemies As EnemyListNode
Public EnemyCount As Long

Public Sub Insert(NewEnemy As IEnemy)
    Dim NewNode As New EnemyListNode
    Set NewNode.Enemy = NewEnemy
    Set NewNode.Nx = Enemies
    Set Enemies = NewNode
End Sub

Public Sub UpdateAll(ByVal TimeDiff As Double)
    Dim Fore As New EnemyListNode
    Set Fore.Nx = Enemies
    Dim It As EnemyListNode
    Set It = Fore
    Dim EnemyCounter As Long
    
    Do While Not It.Nx Is Nothing
        It.Nx.Enemy.Update TimeDiff
        Dim Pos As Point
        Set Pos = It.Nx.Enemy.GetPosition
        If Not It.Nx.Enemy.IsAlive Then
           Set It.Nx = It.Nx.Nx
        Else
            EnemyCounter = EnemyCounter + 1
        End If
        Set It = It.Nx
        If It Is Nothing Then Exit Do
    Loop
    Set Enemies = Fore.Nx
    EnemyCount = EnemyCounter
End Sub

Public Sub Clear()
    Set Enemies = Nothing
End Sub


Private Sub Class_Initialize()
    Set Enemies = Nothing
End Sub

