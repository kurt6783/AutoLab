Public Class Seat
    Dim myBottleCount As Integer = 0
    Dim myBottleFinish As Integer = 0
    Dim myBottles() As Bottle
    Public Sub ScanBottle(ByVal PLC_IP As String, ByVal PLC_Port As Integer, ByVal Seat As Integer)
        Dim PLC As PLC = New PLC
        PLC.CreatConnect(PLC_IP, PLC_Port)
        Dim BottlesHex As String = PLC.Read(20000 + Seat, 1)
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
                myBottles(Existence).Name = 101 + (Seat - 1) * 12 + i
                Existence = Existence + 1
            End If
        Next
        myBottleCount = myBottles.Count
        PLC.Dispose()
        PLC = Nothing
    End Sub
    Public ReadOnly Property BottleCount()
        Get
            Return myBottleCount
        End Get
    End Property
    Public ReadOnly Property Bottles(ByVal Index As Integer)
        Get
            Return myBottles(Index)
        End Get
    End Property
    Public Property BottleFinish()
        Get
            Return myBottleFinish
        End Get
        Set(ByVal value)
            myBottleFinish = value
        End Set
    End Property
End Class


