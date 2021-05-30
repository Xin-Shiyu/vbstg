VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BulletPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Bullets As BulletListNode
Public BulletCount As Long

Public Sub Insert(NewBullet As IBullet)
    Dim NewNode As New BulletListNode
    Set NewNode.Bullet = NewBullet
    Set NewNode.Nx = Bullets
    Set Bullets = NewNode
End Sub

Public Sub UpdateAll(ByVal TimeDiff As Double)
    Dim Fore As New BulletListNode
    Set Fore.Nx = Bullets
    Dim It As BulletListNode
    Set It = Fore
    Dim BulletCounter As Long
    
    Do While Not It.Nx Is Nothing
        It.Nx.Bullet.Update TimeDiff
        Dim Pos As Point
        Set Pos = It.Nx.Bullet.GetPosition
        If Not It.Nx.Bullet.IsAlive Then
           Set It.Nx = It.Nx.Nx
        Else
            BulletCounter = BulletCounter + 1
        End If
        Set It = It.Nx
        If It Is Nothing Then Exit Do
    Loop
    Set Bullets = Fore.Nx
    BulletCount = BulletCounter
End Sub

Public Function MinDistance(ByVal X As Double, ByVal Y As Double) As Double
    Dim It As BulletListNode
    Set It = Bullets
    MinDistance = GameWidth + GameHeight
    Do While Not It Is Nothing
        Dim Pos As Point
        Set Pos = It.Bullet.GetPosition
        Dim NewDistance As Double
        NewDistance = Abs(Pos.X - X) + Abs(Pos.Y - Y) - It.Bullet.GetSize + 2 ' To make game easier
        If NewDistance < MinDistance Then MinDistance = NewDistance
        Set It = It.Nx
    Loop
End Function

Public Sub Clear()
    Set Bullets = Nothing
End Sub


Private Sub Class_Initialize()
    Set Bullets = Nothing
End Sub
