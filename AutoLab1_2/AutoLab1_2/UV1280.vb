Public Class UV1280
    Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Int32, ByVal dwShareMode As Int32, ByVal lpSecurityAttributes As IntPtr, ByVal dwCreationDisposition As Int32, ByVal dwFlagsAndAttributes As Int32, ByVal hTemplateFile As IntPtr) As IntPtr
    Public Declare Auto Function WriteFile Lib "kernel32.dll" (ByVal hFile As IntPtr, ByVal lpBuffer As Byte(), ByVal nNumberOfBytesToWrite As Int32, ByRef lpNumberOfBytesWritten As Int32, ByVal lpOverlapped As IntPtr) As Boolean
    Public Declare Auto Function ReadFile Lib "kernel32.dll" (ByVal hFile As IntPtr, ByVal lpBuffer As Byte(), ByVal nNumberOfBytesToRead As Int32, ByRef lpNumberOfBytesRead As Int32, ByVal lpOverlapped As IntPtr) As Boolean
    Public Declare Auto Function CloseHandle Lib "kernel32.dll" (ByVal hObject As IntPtr) As Boolean
    Dim hCOM As IntPtr
    Dim strComPort As String = "COM5"
    Dim GENERIC_READ = &H80000000
    Dim GENERIC_WRITE = &H40000000
    Dim OPEN_EXISTING As Int32 = 3
    Dim FILE_ATTRIBUTE_NORMAL As Int32 = &H80
    Dim bRet As Boolean
    Dim bybuffer() As Byte
    Dim nBytesWritten As Int32
    Dim nBytesRead As Int32
    Dim AbsorbData As String
    Sub SetPort(ByVal Port As String)
        strComPort = Port
    End Sub
    Sub SetWaveLength(ByVal WaveLength As String)
        Dim Data(1) As String
        Data(0) = "設定波長"
        Data(1) = WaveLength
        ProtocolA(Data)
    End Sub
    Sub SetScanSpeed(ByVal ScanSpeed As String)
        Dim Data(1) As String
        Data(0) = "掃瞄速度"
        Data(1) = ScanSpeed
        ProtocolA(Data)
    End Sub
    Sub SetMode(ByVal Mode As String)
        Dim Data(1) As String
        Data(0) = "測定模式"
        Data(1) = Mode
        ProtocolA(Data)
    End Sub
    Sub Reset()
        Dim Data(0) As String
        Data(0) = "自動歸零"
        ProtocolA(Data)
    End Sub
    Sub Motor()
        Dim Data(0) As String
        Data(0) = "注射器控制"
        ProtocolA(Data)
    End Sub
    Function GetAbsorb()
        Dim Data(0) As String
        Data(0) = "數據輸出"
        Return protocolB(Data)
    End Function
    Sub ProtocolA(ByVal Item() As String)
        'PC:建立連線
        hCOM = CreateFile(strComPort, GENERIC_READ + GENERIC_WRITE, 0, IntPtr.Zero, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero)
        'PC:ENQ
        ReDim bybuffer(0)
        bybuffer(0) = 5
        bRet = WriteFile(hCOM, bybuffer, bybuffer.Length, nBytesWritten, IntPtr.Zero)
        'UV:ACK
        ReDim bybuffer(0)
        bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
        Do
            Select Case bybuffer(0)
                Case 6
                    Exit Do
                Case 27
                    MsgBox("UV無法控制")
                    Exit Sub
                Case Else
                    ReDim bybuffer(0)
                    bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
            End Select
        Loop
        'PC:     指令(+NUL)      
        Dim Code As String = Nothing
        Select Case Item(0)
            Case "設定波長"
                Code = "w" & Item(1).ToString & Chr(&H0)
            Case "掃瞄速度"
                Code = "j" & Item(1).ToString & Chr(&H0)
            Case "測定模式"
                Code = "v" & Item(1).ToString & Chr(&H0)
            Case "注射器控制"
                Code = "o" & Chr(&H0)
            Case "自動歸零"
                Code = "x" & Chr(&H0)
        End Select
        bybuffer = System.Text.Encoding.Default.GetBytes(Code)
        bRet = WriteFile(hCOM, bybuffer, bybuffer.Length, nBytesWritten, IntPtr.Zero)
        'UV:ACK
        ReDim bybuffer(0)
        bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
        Do
            Select Case bybuffer(0)
                Case 6
                    Exit Do
                Case 27
                    MsgBox("UV無法控制")
                    Exit Sub
                Case 21
                    bybuffer = System.Text.Encoding.Default.GetBytes(Code)
                    bRet = WriteFile(hCOM, bybuffer, bybuffer.Length, nBytesWritten, IntPtr.Zero)
                    ReDim bybuffer(0)
                    bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
                Case Else
                    ReDim bybuffer(0)
                    bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
            End Select
        Loop
        'UV:EOT
        ReDim bybuffer(0)
        bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
        Do
            Select Case bybuffer(0)
                Case 4
                    Exit Do
                Case 27
                    MsgBox("UV無法控制")
                    Exit Sub
                Case Else
                    ReDim bybuffer(0)
                    bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
            End Select
        Loop
        'PC:ACK
        ReDim bybuffer(0)
        bybuffer(0) = 6
        bRet = WriteFile(hCOM, bybuffer, bybuffer.Length, nBytesWritten, IntPtr.Zero)
        'PC:結束連線
        CloseHandle(hCOM)
    End Sub
    Function protocolB(ByVal Item() As String)
        'PC:建立連線        
        hCOM = CreateFile(strComPort, GENERIC_READ + GENERIC_WRITE, 0, IntPtr.Zero, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero)
        'PC:ENQ
        ReDim bybuffer(0)
        bybuffer(0) = 5
        bRet = WriteFile(hCOM, bybuffer, bybuffer.Length, nBytesWritten, IntPtr.Zero)
        'UV:ACK
        ReDim bybuffer(0)
        bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
        Do
            Select Case bybuffer(0)
                Case 6
                    Exit Do
                Case Else
                    ReDim bybuffer(0)
                    bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
            End Select
        Loop
        'PC:     指令+NUL      
        Dim Code As String = Nothing
        Select Case Item(0)
            Case "數據輸出"
                Code = "d" & Chr(&H0)
        End Select
        bybuffer = System.Text.Encoding.Default.GetBytes(Code)
        bRet = WriteFile(hCOM, bybuffer, bybuffer.Length, nBytesWritten, IntPtr.Zero)
        'UV:ACK
        ReDim bybuffer(0)
        bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
        Do
            Select Case bybuffer(0)
                Case 6
                    Exit Do
                Case Else
                    ReDim bybuffer(0)
                    bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
            End Select
        Loop
        'UV:ENQ
        ReDim bybuffer(0)
        bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
        Do
            Select Case bybuffer(0)
                Case 5
                    Exit Do
                Case Else
                    ReDim bybuffer(0)
                    bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
            End Select
        Loop
        'PC:ACK
        ReDim bybuffer(0)
        bybuffer(0) = 6
        bRet = WriteFile(hCOM, bybuffer, bybuffer.Length, nBytesWritten, IntPtr.Zero)
        'UV:數據+NUL
        ReDim bybuffer(0)
        Dim Receive As String = Nothing
        bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
        Do Until bybuffer(0) <> 0
            Threading.Thread.Sleep(100)
            bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
        Loop

        Do Until bybuffer(0) = 0
            Receive = Receive & bybuffer(0).ToString & ","
            bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
        Loop
        'PC:ACK
        ReDim bybuffer(0)
        bybuffer(0) = 6
        bRet = WriteFile(hCOM, bybuffer, bybuffer.Length, nBytesWritten, IntPtr.Zero)
        'UV:EOT
        ReDim bybuffer(0)
        bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
        Do
            Select Case bybuffer(0)
                Case 4
                    Exit Do
                Case Else
                    ReDim bybuffer(0)
                    bRet = ReadFile(hCOM, bybuffer, 1, nBytesRead, IntPtr.Zero)
            End Select
        Loop
        'PC:ACK
        ReDim bybuffer(0)
        bybuffer(0) = 6
        bRet = WriteFile(hCOM, bybuffer, bybuffer.Length, nBytesWritten, IntPtr.Zero)
        'PC:結束連線
        CloseHandle(hCOM)
        Dim Result() As String
        Result = Receive.Split(",")
        Dim Absorb As String = Nothing
        For i = 1 To Result.Count - 2
            Absorb = Absorb & Chr(Result(i))
        Next
        Return Absorb
    End Function
End Class