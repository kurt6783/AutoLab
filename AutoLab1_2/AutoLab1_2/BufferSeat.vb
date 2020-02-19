Public Class BufferSeat
    Dim mybottleCount As Integer = 0
    Dim myBottles() As Bottle
    Public Sub ScanBottle(ByVal PLC_IP As String, ByVal PLC_Port As Integer)
        Dim PLC As PLC = New PLC
        PLC.CreatConnect(PLC_IP, PLC_Port)
        Dim BottlesHex As String = PLC.Read(20004, 1)
        Dim BottlesDEC As Integer = Convert.ToInt32(BottlesHex, 16)
        Dim BottlesB As String = Convert.ToString(BottlesDEC, 2).PadLeft(12, "0")
        Dim Bottles(BottlesB.Length - 1) As String
        For i = 0 To Bottles.Count - 1
            Bottles(i) = Mid(BottlesB, i + 1, 1)
        Next
        Array.Reverse(Bottles)
        Dim Existence As Integer = 0
        For i = 0 To Bottles.Count - 1
            If Bottles(i) = 1 Then
                Existence = Existence + 1
            End If
        Next
        ReDim myBottles(Existence - 1)
        Existence = 0
        For i = 0 To Bottles.Count - 1
            If Bottles(i) = 1 Then
                myBottles(Existence) = New Bottle
                myBottles(Existence).Name = 601 + i
                If 0 <= i And i <= 10 Then
                    myBottles(Existence).Type = "SGPickling"
                ElseIf 11 <= i And i <= 11 Then
                    myBottles(Existence).Type = "PHSoak"
                End If
                Existence = Existence + 1
            End If
        Next
        mybottleCount = myBottles.Count
        PLC.Dispose()
        PLC = Nothing
    End Sub
    Public ReadOnly Property BottleCount()
        Get
            Return mybottleCount
        End Get
    End Property
    Public ReadOnly Property Bottles(ByVal Index As Integer)
        Get
            Return myBottles(Index)
        End Get
    End Property
End Class
