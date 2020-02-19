Imports System.Net.Sockets
Public Class PLC
    Private Client
    Private Stream As NetworkStream
    Private NodeAddress As String
    Public Sub CreatConnect(ByVal IP As String, ByVal Port As Integer)
        Client = New TcpClient(IP, Port)
        Stream = Client.GetStream
        Stream.ReadTimeout = 1000
        Stream.WriteTimeout = 1000
        Dim Head As String = "46494E530000000C000000000000000000000000"
        Dim Length As Integer = Head.Length / 2 - 1
        Dim SendData(Length) As Byte
        Dim j As Integer = 1
        For i = 0 To Length
            SendData(i) = "&h" & Mid(Head, j, 2)
            j = j + 2
        Next
        Stream.Write(SendData, 0, SendData.Length)
        Dim ReceiveData(24) As Byte
        Stream.Read(ReceiveData, 0, ReceiveData.Length)
        If ReceiveData(11) = 1 Then
            NodeAddress = Microsoft.VisualBasic.Right("0" & Hex(CLng(ReceiveData(19))), 2)
        End If
    End Sub
    Public Function Read(ByVal DM As Integer, ByVal Amount As Integer)
        Dim Head As String = "46494E530000001A0000000200000000"
        Dim ICF As String = "80"
        Dim RSV As String = "00"
        Dim GCT As String = "02"
        Dim DNA As String = "00"
        Dim DA1 As String = "00"
        Dim DA2 As String = "00"
        Dim SNA As String = "00"
        Dim SA1 As String = "00"
        Dim SA2 As String = "00"
        Dim SID As String = "00"
        Dim Length As String = Hex(Amount).ToString.PadLeft(6, "0")
        Dim Text = Hex(Val(DM))
        Select Case Len(Text)
            Case 0 : Text = "0000"
            Case 1 : Text = "000" & Text
            Case 2 : Text = "00" & Text
            Case 3 : Text = "0" & Text
            Case Else : Text = Text
        End Select
        Dim Commend As String = Head & ICF & RSV & GCT & DNA & DA1 & DA2 & SNA & NodeAddress & SA2 & SID & "010182" & Text & Length
        Dim SendData(Commend.Length / 2 - 1) As Byte
        Dim j As Integer = 1
        For i = 0 To Commend.Length / 2 - 1
            SendData(i) = "&h" & Mid(Commend, j, 2)
            j = j + 2
        Next
        Stream.Write(SendData, 0, SendData.Length)
        Dim ReceiveData(29 + 2 * Amount) As Byte
        Stream.Read(ReceiveData, 0, ReceiveData.Length)
        Dim Receive As String = Nothing
        For i = 0 To Amount * 2 - 1
            Receive = Receive & Hex(ReceiveData(30 + i)).ToString.PadLeft(2, "0")
        Next
        Return Receive
    End Function
    Public Sub Write(ByVal DM As Integer, ByVal Data As String)
        Dim Head As String = "46494E530000001C0000000200000000"
        Dim ICF As String = "80"
        Dim RSV As String = "00"
        Dim GCT As String = "02"
        Dim DNA As String = "00"
        Dim DA1 As String = "00"
        Dim DA2 As String = "00"
        Dim SNA As String = "00"
        Dim SA1 As String = "00"
        Dim SA2 As String = "00"
        Dim SID As String = "00"
        Dim Text = Hex(Val(DM))
        Select Case Len(Text)
            Case 0 : Text = "0000"
            Case 1 : Text = "000" & Text
            Case 2 : Text = "00" & Text
            Case 3 : Text = "0" & Text
            Case Else : Text = Text
        End Select
        Dim Commend As String = Head & ICF & RSV & GCT & DNA & DA1 & DA2 & SNA & NodeAddress & SA2 & SID & "010282" & Text & Format(1, "000000") & Data.PadLeft(4, "0")
        Dim SendData(Commend.Length / 2 - 1) As Byte
        Dim j As Integer = 1
        For i = 0 To Commend.Length / 2 - 1
            SendData(i) = "&h" & Mid(Commend, j, 2)
            j = j + 2
        Next
        Stream.Write(SendData, 0, SendData.Length)
        Dim ReceiveData(100) As Byte
        Stream.Read(ReceiveData, 0, 100)
    End Sub
    Sub Dispose()
        Stream.Dispose()
    End Sub
End Class
