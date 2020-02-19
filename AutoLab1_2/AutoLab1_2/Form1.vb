Imports System.IO.Ports
Imports Newtonsoft.Json.Linq
Imports System.Threading

Public Class Form1
    '-------------------------------------------------------------------DataBase
    Dim DataBase As DataBase = New DataBase
    '-------------------------------------------------------------------Platform
    Dim UVPlatform As Platform = New Platform
    Dim PHPlatform As Platform = New Platform
    Dim SGPlatform As Platform = New Platform
    '-------------------------------------------------------------------DataTable
    Dim MessageDataTable As DataTable = New DataTable
    '-------------------------------------------------------------------Thread
    Dim SystemFlowThread As Threading.Thread
    Dim SeatThread As Threading.Thread
    Dim SGThread As Threading.Thread
    Dim SGPicklingThread As Threading.Thread
    Dim PHThread As Threading.Thread
    Dim PHSoakThread As Threading.Thread
    Dim UVThread As Threading.Thread
    '-------------------------------------------------------------------PLC
    Dim PLC_IP As String = "192.168.50.209"
    Dim PLC_Port As Integer = 9600
    Dim SystemPLC As PLC
    Dim RobotPLC As PLC
    Dim PAUSEPLC As PLC
    Dim SeatPLC As PLC
    Dim SGFlowPLC As PLC
    Dim SGPicklingFlowPLC As PLC
    Dim PHFlowPLC As PLC
    Dim PHSoakFlowPLC As PLC
    Dim UVFlowPLC As PLC
    '-------------------------------------------------------------------Queue
    Dim Newly As Queue = New Queue
    Dim SGMeasure As Queue = New Queue
    Dim SGQueue As Queue = New Queue
    Dim SGPickling As Queue = New Queue
    Dim PHMeasure As Queue = New Queue
    Dim PHQueue As Queue = New Queue
    Dim PHSoak As Queue = New Queue
    Dim UVMeasure As Queue = New Queue
    Dim UVQueue As Queue = New Queue
    Dim Finish As Queue = New Queue
    Dim Seat As Queue = New Queue
    '-------------------------------------------------------------------Seat
    Dim Seat1 As Seat
    Dim Seat2 As Seat
    Dim Seat3 As Seat
    '-------------------------------------------------------------------BufferSeat
    Dim BufferSeat As BufferSeat
    '-------------------------------------------------------------------Scanner
    Dim WithEvents ScanTimer As New System.Timers.Timer
    Dim WithEvents Scanner As New System.IO.Ports.SerialPort
    Dim ScanTimer_Time As Integer = 10000
    Dim Scanner_Port As String = "COM8"
    Dim QRCodeStatus As Integer = 0
    Dim QRCode As String = Nothing
    '-------------------------------------------------------------------WashTimer
    Dim WithEvents WashTimer As New System.Timers.Timer
    Dim WashTimer_Time As Integer = 30000
    '-------------------------------------------------------------------FormEvent
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            '---------------------------------------------------------------------------------------------------------系統初始化
            SystemPLC = New PLC
            SystemPLC.CreatConnect(PLC_IP, PLC_Port)
            SystemPLC.Write(0, "0000") '---------------手臂起始位置
            SystemPLC.Write(1, "0000") '---------------手臂目標位置
            SystemPLC.Write(100, "0000") '------------SG狀態歸零
            SystemPLC.Write(200, "0000") '------------PH狀態歸零
            SystemPLC.Write(300, "0000") '------------UV狀態歸零
            SystemPLC.Write(9998, "0002")   '---------三色燈狀態待機
            '-------------------------------------------------------------------------------------------------------啟用執行緒
            Form1.CheckForIllegalCrossThreadCalls = False
            SeatThread = New System.Threading.Thread(AddressOf SeatStatus)
            SeatThread.IsBackground = True
            SeatThread.Start()
            '------------------------------------------------------------------------------------------------------設定訊息表
            MessageDataTable.Columns.Add("Time")
            MessageDataTable.Columns.Add("Message")
            MessageDataTable.DefaultView.Sort = "Time DESC"
            DataGridView1.DataSource = MessageDataTable
            MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "System Open。")
            ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "System Open。")
            '-----------------------------------------------------------------------------------------------------新增工作平台
            For i = 0 To ListBox9.Items.Count - 1
                Select Case ListBox9.Items(i)
                    Case "UV"
                        UVPlatform.Status = "StandBy"
                    Case "PH"
                        PHPlatform.Status = "StandBy"
                    Case "SG"
                        SGPlatform.Status = "StandBy"
                End Select
            Next
            BTNSTART.Enabled = False
            BTNPASUE.Enabled = False
            BTNSEAT1.Enabled = False
            BTNSEAT2.Enabled = False
            BTNSEAT3.Enabled = False
            BTNREFRESH.Enabled = False
            BTNENABLE.Enabled = False
            BTNDISABLE.Enabled = False
            BTNSGDOWN.Enabled = False
            BTNSGWASH.Enabled = False
            BTNPHDOWN.Enabled = False
            BTNPHWASH.Enabled = False
            BTNUVDOWN.Enabled = False
            BTNUVWASH.Enabled = False
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("FORMLOAD Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        MsgBox("Please confirm that there is no content on each platform。" & vbCrLf & "請淨空各工作平台。")
    End Sub
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "System Close。")
        Dim a() = Process.GetProcessesByName("AutoLab1_2")
        If a.Count > 0 Then
            For i = 0 To a.Count - 1
                a(i).Kill()
            Next
        End If
    End Sub
    Private Sub BTNSTART_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSTART.Click
        Try
            If SGPickling.Count = 0 And PHSoak.Count = 0 Then
                MsgBox("Please refresh the BufferSeat and restart the system")
                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "System Stop。")
                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "System Stop。")
            Else
                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "System Start。")
                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "System Start。")
                SystemFlowThread = New System.Threading.Thread(AddressOf SystemFlow)
                SystemFlowThread.IsBackground = True
                SystemFlowThread.Start()
                BTNSTART.Enabled = False
                BTNPASUE.Enabled = True
                BTNSEAT1.Enabled = False
                BTNSEAT2.Enabled = False
                BTNSEAT3.Enabled = False
                BTNREFRESH.Enabled = False
                BTNUVDOWN.Enabled = False
                BTNUVWASH.Enabled = False
                BTNPHDOWN.Enabled = False
                BTNPHWASH.Enabled = False
                BTNSGDOWN.Enabled = False
                BTNSGWASH.Enabled = False
                BTNENABLE.Enabled = False
                BTNDISABLE.Enabled = False
                SystemPLC.Write(9998, "0001")
            End If
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("BTNSTART Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Private Sub BTNPAUSE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPASUE.Click
        Try
            PAUSEPLC = New PLC
            PAUSEPLC.CreatConnect(PLC_IP, PLC_Port)
            Select Case BTNPASUE.Text
                Case "PAUSE"
                    MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "System PAUSE。")
                    ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "System PAUSE。")
                    Do Until PAUSEPLC.Read(0, 1) = "0000" And PAUSEPLC.Read(1, 1) = "0000"
                        Threading.Thread.Sleep(100)
                    Loop
                    SystemFlowThread.Suspend()
                    Threading.Thread.Sleep(5000)
                    If Seat.Contains(1) = False Then
                        BTNSEAT1.Enabled = True
                    End If
                    If Seat.Contains(2) = False Then
                        BTNSEAT2.Enabled = True
                    End If
                    If Seat.Contains(3) = False Then
                        BTNSEAT3.Enabled = True
                    End If
                    BTNREFRESH.Enabled = True
                    BTNPASUE.Text = "RESUME"
                    PAUSEPLC.Write(9998, "0002")
                Case "RESUME"
                    MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "System RESUME。")
                    ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "System RESUME。")
                    SystemFlowThread.Resume()
                    BTNPASUE.Text = "PAUSE"
                    BTNSEAT1.Enabled = False
                    BTNSEAT2.Enabled = False
                    BTNSEAT3.Enabled = False
                    BTNREFRESH.Enabled = False
                    PAUSEPLC.Write(9998, "0001")
            End Select
            PAUSEPLC.Dispose()
            PAUSEPLC = Nothing
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("BTNPAUSE Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Private Sub BTNSTOP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If MsgBox("確認停止?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "System STOP。")
                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "System STOP。")
                If SystemFlowThread.ThreadState = ThreadState.Suspended Then
                    SystemFlowThread.Resume()
                End If
                SystemFlowThread.Abort()
                SystemFlowThread = Nothing
                If UVThread IsNot Nothing Then
                    If UVThread.ThreadState = Threading.ThreadState.WaitSleepJoin Or UVThread.ThreadState = Threading.ThreadState.Running Then
                        UVThread.Abort()
                        UVThread = Nothing
                    End If
                End If
                If PHThread IsNot Nothing Then
                    If PHThread.ThreadState = Threading.ThreadState.WaitSleepJoin Or PHThread.ThreadState = Threading.ThreadState.Running Then
                        PHThread.Abort()
                        PHThread = Nothing
                    End If
                End If
                If SGThread IsNot Nothing Then
                    If SGThread.ThreadState = Threading.ThreadState.WaitSleepJoin Or SGThread.ThreadState = Threading.ThreadState.Running Then
                        SGThread.Abort()
                        SGThread = Nothing
                    End If
                End If
                For i = 1 To 36
                    SeatStatus(i, Color.Black)
                Next
                PHPlatform.Status = "StandBy"
                SGPlatform.Status = "StandBy"
                UVPlatform.Status = "StandBy"
                Newly.Clear()
                UVMeasure.Clear()
                UVQueue.Clear()
                PHMeasure.Clear()
                PHQueue.Clear()
                SGMeasure.Clear()
                SGQueue.Clear()
                Finish.Clear()
                SGPickling.Clear()
                PHSoak.Clear()
                Seat.Clear()
                ShowView()
                Seat1 = Nothing
                Seat2 = Nothing
                Seat3 = Nothing
                BufferSeat = Nothing
                BTNSTART.Enabled = True
                BTNSEAT1.Enabled = True
                BTNSEAT2.Enabled = True
                BTNSEAT3.Enabled = True
                BTNREFRESH.Enabled = True
                BTNUVDOWN.Enabled = True
                BTNUVWASH.Enabled = True
                BTNPHDOWN.Enabled = True
                BTNPHWASH.Enabled = True
                BTNSGDOWN.Enabled = True
                BTNSGWASH.Enabled = True
                BTNENABLE.Enabled = True
                BTNDISABLE.Enabled = True
                BTNPASUE.Enabled = False
                BTNPASUE.Text = "PAUSE"
                BTNENABLE.Enabled = True
                BTNDISABLE.Enabled = True
                SystemPLC.Write(0, "0000")
                SystemPLC.Write(1, "0000")
                SystemPLC.Write(100, "0000")
                SystemPLC.Write(200, "0000")
                SystemPLC.Write(300, "0000")
                SystemPLC.Write(9998, "0002")
                MsgBox("Please confirm that there is no content on each platform。" & vbCrLf & "請淨空各工作平台。")
            End If
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("BTNSTOP Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Private Sub BTNSEAT1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSEAT1.Click
        Try
            If BTNSEAT1.Text = "Seat1 Finish" Then
                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Refresh Seat1。")
                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "Refresh Seat1。")
                BTNSEAT1.Text = "Join Seat1"
                BTNSEAT1.ForeColor = Color.Black
                For i = 0 To 11
                    SeatStatus(i + 1, Color.Black)
                Next
            Else
                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Seat1 Join。")
                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "Seat1 Join。")
                DataBase.ConnectDB()
                Seat1 = New Seat
                Seat1.ScanBottle(PLC_IP, PLC_Port, 1)
                For i = 0 To Seat1.BottleCount - 1
                    Newly.Enqueue(Seat1.Bottles(i))
                Next
                Seat.Enqueue(1)
                CheckSeat()
                ShowView()
                BTNSEAT1.Enabled = False
            End If
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("BTNSEAT1 Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Private Sub BTNSEAT2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSEAT2.Click
        Try
            If BTNSEAT2.Text = "Seat2 Finish" Then
                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Refresh Seat2。")
                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "Refresh Seat2。")
                BTNSEAT2.Text = "Join Seat2"
                BTNSEAT2.ForeColor = Color.Black
                For i = 0 To 11
                    SeatStatus(i + 13, Color.Black)
                Next
            Else
                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Seat2 Join。")
                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "Seat2 Join。")
                DataBase.ConnectDB()
                Seat2 = New Seat
                Seat2.ScanBottle(PLC_IP, PLC_Port, 2)
                For i = 0 To Seat2.BottleCount - 1
                    Newly.Enqueue(Seat2.Bottles(i))
                Next
                Seat.Enqueue(2)
                CheckSeat()
                ShowView()
                BTNSEAT2.Enabled = False
            End If
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("BTNSEAT2 Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Private Sub BTNSEAT3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSEAT3.Click
        Try
            If BTNSEAT3.Text = "Seat3 Finish" Then
                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Refresh Seat3。")
                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "Refresh Seat3。")
                BTNSEAT3.Text = "Join Seat3"
                BTNSEAT3.ForeColor = Color.Black
                For i = 0 To 11
                    SeatStatus(i + 25, Color.Black)
                Next
            Else
                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Seat3 Join。")
                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "Seat3 Join。")
                DataBase.ConnectDB()
                Seat3 = New Seat
                Seat3.ScanBottle(PLC_IP, PLC_Port, 3)
                For i = 0 To Seat3.BottleCount - 1
                    Newly.Enqueue(Seat3.Bottles(i))
                Next
                Seat.Enqueue(3)
                CheckSeat()
                ShowView()
                BTNSEAT3.Enabled = False
            End If
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("BTNSEAT3 Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Private Sub BTNUVDOWN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNUVDOWN.Click
        If BTNUVDOWN.Text = "UVDown" Then
            SystemPLC.Write(300, "FFFF")
            BTNUVDOWN.Text = "UVUp"
        ElseIf BTNUVDOWN.Text = "UVUp" Then
            SystemPLC.Write(300, "0000")
            BTNUVDOWN.Text = "UVDown"
        End If
    End Sub
    Private Sub BTNPHDOWN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPHDOWN.Click
        If BTNPHDOWN.Text = "PHDown" Then
            SystemPLC.Write(200, "FFFF")
            BTNPHDOWN.Text = "PHUp"
        ElseIf BTNPHDOWN.Text = "PHUp" Then
            SystemPLC.Write(200, "0000")
            BTNPHDOWN.Text = "PHDown"
        End If
    End Sub
    Private Sub BTNSGDOWN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSGDOWN.Click
        If BTNSGDOWN.Text = "SGDown" Then
            SystemPLC.Write(100, "FFFF")
            BTNSGDOWN.Text = "SGUp"
        ElseIf BTNSGDOWN.Text = "SGUp" Then
            SystemPLC.Write(100, "0000")
            BTNSGDOWN.Text = "SGDown"
        End If
    End Sub
    Private Sub BTNUVWASH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNUVWASH.Click
        If BTNUVWASH.Text = "UVWash" Then
            SystemPLC.Write(300, "FFFE")
            BTNUVWASH.Text = "UVWashStop"
        ElseIf BTNUVWASH.Text = "UVWashStop" Then
            SystemPLC.Write(300, "0000")
            BTNUVWASH.Text = "UVWash"
        End If
    End Sub

    Private Sub BTNPHWASH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPHWASH.Click
        If BTNPHWASH.Text = "PHWash" Then
            SystemPLC.Write(200, "FFFE")
            BTNPHWASH.Text = "PHWashStop"
        ElseIf BTNPHWASH.Text = "PHWashStop" Then
            SystemPLC.Write(200, "0000")
            BTNPHWASH.Text = "PHWash"
        End If
    End Sub

    Private Sub BTNSGWASH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSGWASH.Click
        If BTNSGWASH.Text = "SGWash" Then
            SystemPLC.Write(100, "FFFE")
            BTNSGWASH.Text = "SGWashStop"
        ElseIf BTNSGWASH.Text = "SGWashStop" Then
            SystemPLC.Write(100, "0000")
            BTNSGWASH.Text = "SGWash"
        End If
    End Sub

    Private Sub BTNALLWASH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNALLWASH.Click
        Try
            MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Wash all platforms。")
            ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "Wash all platforms。")
            SystemPLC.Write(100, "FFFE")
            SystemPLC.Write(200, "FFFE")
            SystemPLC.Write(300, "FFFE")
            WashTimer.Interval = WashTimer_Time
            WashTimer.Start()
            BTNALLWASH.Enabled = False
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("BTNALLWASH Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub

    Sub WashFinished() Handles WashTimer.Elapsed
        Try
            WashTimer.Close()
            SystemPLC.Write(100, "0000")
            SystemPLC.Write(200, "0000")
            SystemPLC.Write(300, "0000")
            BTNSTART.Enabled = True
            BTNSGDOWN.Enabled = True
            BTNSGWASH.Enabled = True
            BTNPHDOWN.Enabled = True
            BTNPHWASH.Enabled = True
            BTNUVDOWN.Enabled = True
            BTNUVWASH.Enabled = True
            BTNENABLE.Enabled = True
            BTNDISABLE.Enabled = True
            BTNREFRESH.Enabled = True
            MsgBox("All platforms had been washed。")
            MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "All platforms had been washed。")
            ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "All platforms had been washed。")
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("WashFinished Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub

    Private Sub BTNREFRESH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNREFRESH.Click
        Try
            MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Buffer Refresh。")
            ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "Buffer Refresh")
            SGPickling.Clear()
            PHSoak.Clear()
            Dim FinishBottle As Bottle
            For i = 0 To Finish.Count - 1
                FinishBottle = Finish.Peek
                If 600 < Val(FinishBottle.Name) And Val(FinishBottle.Name) < 613 Then
                    Finish.Dequeue()
                Else
                    Finish.Enqueue(Finish.Dequeue)
                End If
            Next
            BufferSeat = New BufferSeat
            BufferSeat.ScanBottle(PLC_IP, PLC_Port)
            Dim Bottle As Bottle
            For i = 0 To BufferSeat.BottleCount - 1
                Bottle = BufferSeat.Bottles(i)
                If Bottle.Type = "SGPickling" Then
                    SGPickling.Enqueue(Bottle)
                ElseIf Bottle.Type = "PHSoak" Then
                    PHSoak.Enqueue(Bottle)
                End If
            Next
            ShowView()
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("BTNREFRESH Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Private Sub BTNENABLE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNENABLE.Click
        Try
            If ListBox10.SelectedItem IsNot Nothing Then
                ListBox9.Items.Add(ListBox10.SelectedItem)
                Select Case ListBox10.SelectedItem
                    Case "UV"
                        UVPlatform.Status = "StandBy"
                    Case "PH"
                        PHPlatform.Status = "StandBy"
                    Case "SG"
                        SGPlatform.Status = "StandBy"
                End Select
                ListBox10.Items.Remove(ListBox10.SelectedItem)
            End If
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("BTNENABLE Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Private Sub BTNDISABLE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNDISABLE.Click
        Try
            If ListBox9.SelectedItem IsNot Nothing Then
                ListBox10.Items.Add(ListBox9.SelectedItem)
                Select Case ListBox9.SelectedItem
                    Case "UV"
                        UVPlatform.Status = "ShutDown"
                    Case "PH"
                        PHPlatform.Status = "ShutDown"
                    Case "SG"
                        SGPlatform.Status = "ShutDown"
                End Select
                ListBox9.Items.Remove(ListBox9.SelectedItem)
            End If
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("BTNDISABLE Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    '--------------------------------------------------------------------New Function
    Sub SystemFlow()
        Try
            Do
                If PHPlatform.Status = "Remove" Then                                                                            '--------PH出
                    PHOut()
                ElseIf UVPlatform.Status = "Remove" Then                                                                        '--------UV出
                    UVOut()
                ElseIf SGPlatform.Status = "Remove" Then                                                                        '--------SG出
                    SGOut()
                ElseIf PHPlatform.Status = "StandBy" And PHQueue.Count = 0 And Newly.Count = 0 Then                             '--------PH待機進      (Thread)   
                    PHDomancyIn()
                ElseIf PHPlatform.Status = "StandBy" And PHQueue.Count > 0 Then                                                 '--------PH進          (Thread)
                    PHIn()
                ElseIf UVPlatform.Status = "StandBy" And UVQueue.Count > 0 Then                                                 '--------UV進          (Thread)
                    UVIn()
                ElseIf SGPlatform.Status = "StandBy" And SGQueue.Count > 0 Then                                                 '--------SG進          (Thread)
                    SGIn()
                ElseIf SGPlatform.Status = "Pickling" Then                                                                      '--------SG酸洗進      (Thread)
                    SGPicklingIn()
                ElseIf Newly.Count > 0 Then                                                                                     '--------掃新瓶
                    ScanNew()
                Else
                    Threading.Thread.Sleep(100)
                End If
                Threading.Thread.Sleep(1000)
                CheckSeat()
                ShowView()
            Loop
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("SystemFlow Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub

    Sub PHOut()                        '--------PH出
        Try
            Dim PHBottle As Bottle = PHMeasure.Peek
            Robot("0502", PHBottle.Name.ToString.PadLeft(4, "0"))
            If PHBottle.Name = "612" Then
                PHSoak.Enqueue(PHMeasure.Dequeue)
            Else
                Finish.Enqueue(PHMeasure.Dequeue)
                BottleFinish(PHBottle)
                SeatStatus(Val(PHBottle.Name) - 100, Color.Blue)
                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & PHBottle.Name & "-Finish")
                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & PHBottle.Name & "-Finish")
            End If
            PHPlatform.Status = "Washing"
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("PHOut Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub UVOut()                       '--------UV出
        Try
            Dim UVBottle As Bottle = UVMeasure.Peek
            Robot("0503", UVBottle.Name.ToString.PadLeft(4, "0"))
            BottleFinish(UVBottle)
            Finish.Enqueue(UVMeasure.Dequeue)
            SeatStatus(Val(UVBottle.Name) - 100, Color.Blue)
            MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & UVBottle.Name & "-Finish")
            ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & UVBottle.Name & "-Finish")
            UVPlatform.Status = "Washing"
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("UVOut Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub SGOut()                       '--------SG出
        Try
            Dim SGBottle As Bottle = SGMeasure.Peek
            Robot("0501", SGBottle.Name.ToString.PadLeft(4, "0"))
            If 600 <= SGBottle.Name And SGBottle.Name <= 612 Then
                SGMeasure.Dequeue()
                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & SGBottle.Name & "-Finish")
                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & SGBottle.Name & "-Finish")
            Else
                Finish.Enqueue(SGMeasure.Dequeue)
                BottleFinish(SGBottle)
                SeatStatus(Val(SGBottle.Name) - 100, Color.Blue)
                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & SGBottle.Name & "-Finish")
                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & SGBottle.Name & "-Finish")
            End If
            SGPlatform.Status = "Washing"
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("SGOut Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub PHDomancyIn()          '--------PH待機入      (Thread)   
        Try
            Dim PHSoakBottle As Bottle = PHSoak.Peek
            Robot(PHSoakBottle.Name.ToString.PadLeft(4, "0"), "0400")
            Robot("0400", "0502")
            PHMeasure.Enqueue(PHSoak.Dequeue())        '
            PHSoakThread = New System.Threading.Thread(AddressOf PHSoakFlow)
            PHSoakThread.IsBackground = False
            PHSoakThread.Start()
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("PHDomancyThread")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub PHIn()                          '--------PH入          (Thread)
        Try
            Dim PHBottle As Bottle = PHQueue.Peek
            Robot(PHBottle.Name.ToString.PadLeft(4, "0"), "0400")
            Do Until SystemPLC.Read(3, 1) = "0001" Or SystemPLC.Read(3, 1) = "0002"
                Threading.Thread.Sleep(100)
            Loop
            Select Case SystemPLC.Read(3, 1)
                Case "0001"   '------------------------------------------開蓋成功
                    Robot("0400", "0502")
                    PHMeasure.Enqueue(PHQueue.Dequeue())
                    PHThread = New System.Threading.Thread(AddressOf PHFlow)
                    PHThread.IsBackground = False
                    PHThread.Start()
                Case "0002" '-------------------------------------------開蓋失敗
                    Robot("400", PHBottle.Name.ToString.PadLeft(4, "0"))
                    BottleFinish(PHBottle)
                    Finish.Enqueue(PHQueue.Dequeue())
                    SeatStatus(Val(PHBottle.Name) - 100, Color.Red)
                    MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & PHBottle.Name & "-OpenCan Failed")
                    ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & PHBottle.Name & "-OpenCan Failed")
            End Select
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("PHThread")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub UVIn()                          '--------UV入          (Thread)
        Try
            Dim UVBottle As Bottle = UVQueue.Peek
            Robot(UVBottle.Name.ToString.PadLeft(4, "0"), "0400")
            Do Until SystemPLC.Read(3, 1) = "0001" Or SystemPLC.Read(3, 1) = "0002"
                Threading.Thread.Sleep(100)
            Loop
            Select Case SystemPLC.Read(3, 1)
                Case "0001"   '------------------------------------------開蓋成功
                    Robot("0400", "0503")
                    UVMeasure.Enqueue(UVQueue.Dequeue())
                    UVThread = New System.Threading.Thread(AddressOf UVFlow)
                    UVThread.IsBackground = False
                    UVThread.Start()
                Case "0002" '-------------------------------------------開蓋失敗
                    Robot("400", UVBottle.Name.ToString.PadLeft(4, "0"))
                    BottleFinish(UVBottle)
                    Finish.Enqueue(UVQueue.Dequeue())
                    SeatStatus(Val(UVBottle.Name) - 100, Color.Red)
                    MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & UVBottle.Name & "-OpenCan Failed")
                    ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & UVBottle.Name & "-OpenCan Failed")
            End Select
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("UVThread")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub SGIn()                          '--------SG入          (Thread)
        Try
            Dim SGBottle As Bottle = SGQueue.Peek
            Dim BottleFormula As JArray
            BottleFormula = JArray.Parse(SGBottle.Formula)
            Dim BottleResult = From objs In BottleFormula.Values(Of JObject)() Where objs("actVal").ToString() Select objs
            If BottleResult.Single.Item("actVal").ToString = 2 And SGPickling.Count = 0 Then    '-------------無足夠酸洗液可使用
                BottleFinish(SGBottle)
                Finish.Enqueue(SGQueue.Dequeue())
                SeatStatus(Val(SGBottle.Name) - 100, Color.Red)
                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & SGBottle.Name & "-No Enough PicklingBottle To Use")
                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & SGBottle.Name & "-No Enough PicklingBottle To Use")
            Else
                Robot(SGBottle.Name.ToString.PadLeft(4, "0"), "0400")
                Do Until SystemPLC.Read(3, 1) = "0001" Or SystemPLC.Read(3, 1) = "0002"
                    Threading.Thread.Sleep(100)
                Loop
                Select Case SystemPLC.Read(3, 1)
                    Case "0001"   '------------------------------------------開蓋成功
                        Robot("0400", "0501")
                        SGMeasure.Enqueue(SGQueue.Dequeue())
                        SGThread = New System.Threading.Thread(AddressOf SGFlow)
                        SGThread.IsBackground = False
                        SGThread.Start()
                    Case "0002" '-------------------------------------------開蓋失敗
                        Robot("400", SGBottle.Name.ToString.PadLeft(4, "0"))
                        BottleFinish(SGBottle)
                        Finish.Enqueue(SGQueue.Dequeue())
                        SeatStatus(Val(SGBottle.Name) - 100, Color.Red)
                        MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & SGBottle.Name & "-OpenCan Failed")
                        ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & SGBottle.Name & "-OpenCan Failed")
                End Select
            End If
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("SGThread")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub SGPicklingIn()            '--------SG酸洗入      (Thread)
        Try
            Dim SGBottle As Bottle = SGPickling.Peek
            Robot(SGBottle.Name.ToString.PadLeft(4, "0"), "0400")
            Do Until SystemPLC.Read(3, 1) = "0001" Or SystemPLC.Read(3, 1) = "0002"
                Threading.Thread.Sleep(100)
            Loop
            Select Case SystemPLC.Read(3, 1)
                Case "0001"   '------------------------------------------開蓋成功
                    Robot("0400", "0501")
                    SGMeasure.Enqueue(SGPickling.Dequeue())
                    SGPicklingThread = New System.Threading.Thread(AddressOf SGPicklingFlow)
                    SGPicklingThread.IsBackground = False
                    SGPicklingThread.Start()
                Case "0002" '-------------------------------------------開蓋失敗
                    Robot("0400", SGBottle.Name.ToString.PadLeft(4, "0"))
                    SGPickling.Dequeue()
                    SeatStatus(Val(SGBottle.Name) - 100, Color.Red)
                    MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & SGBottle.Name & "-OpenCan Failed")
                    ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & SGBottle.Name & "-OpenCan Failed")
                    If SGPickling.Count = 0 Then
                        'BTNPASUE.c
                    End If
            End Select
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("SGPicklingThread")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub ScanNew()                  '--------掃新瓶
        Try
            '----------------------------------------------------送入掃描
            Dim NewlyBottle As Bottle = Newly.Peek
            QRCodeStatus = 0
            SystemPLC.Write(0, NewlyBottle.Name.ToString.PadLeft(4, "0"))
            SystemPLC.Write(1, "0200")
            Do Until SystemPLC.Read(2, 1) = "0200"
                Threading.Thread.Sleep(100)
            Loop
            '------------------------------------------------------開始掃描
            ScanTimer.Interval = ScanTimer_Time
            ScanTimer.Start()
            Scanner.PortName = Scanner_Port
            Scanner.BaudRate = 115200
            Scanner.Parity = Parity.None
            Scanner.DataBits = 8
            Scanner.StopBits = StopBits.One
            Scanner.Open()
            Do Until QRCodeStatus <> 0
                Threading.Thread.Sleep(100)
            Loop
            SystemPLC.Write(3, QRCodeStatus.ToString.PadLeft(4, "0"))
            NewlyBottle.PreprocessDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Select Case QRCodeStatus
                Case 1 '----------掃描超時
                    Robot("0200", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                    Finish.Enqueue(Newly.Dequeue())
                    BottleFinish(NewlyBottle)
                    SeatStatus(Val(NewlyBottle.Name) - 100, Color.Red)
                    MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & NewlyBottle.Name & "-Time Out")
                    ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & NewlyBottle.Name & "-Time Out")
                Case 2 '----------掃描成功
                    NewlyBottle.BottleId = QRCode
                    NewlyBottle.ItemSn = DataBase.itemSn(NewlyBottle.BottleId)
                    NewlyBottle.MethodId = DataBase.methodId(NewlyBottle.BottleId)
                    NewlyBottle.MethodSn = DataBase.methodSn(NewlyBottle.BottleId)
                    NewlyBottle.OnmicSn = DataBase.onmicSn(NewlyBottle.BottleId)
                    NewlyBottle.SetDate = DataBase.setDate(NewlyBottle.BottleId)
                    NewlyBottle.TankSn = DataBase.tankSn(NewlyBottle.BottleId)
                    NewlyBottle.Formula = DataBase.formula(NewlyBottle.BottleId)
                    Select Case NewlyBottle.OnmicSn
                        Case "61"
                            If PHPlatform.Status = "ShutDown" Then '------------------------------------------------------掃描成功、PH平台關閉中
                                Robot("0200", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                                Finish.Enqueue(Newly.Dequeue)
                                BottleFinish(NewlyBottle)
                                SeatStatus(Val(NewlyBottle.Name) - 100, Color.Red)
                                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & NewlyBottle.Name & "-PHPlatform had been ShutDown")
                                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & NewlyBottle.Name & "-Platform had been ShutDown")
                            ElseIf PHMeasure.Count > 0 Then '---------------------------------------------------------掃描成功、PH平台沒空
                                Robot("0200", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                                PHQueue.Enqueue(Newly.Dequeue())
                            ElseIf PHMeasure.Count = 0 And PHPlatform.Status <> "StandBy" Then '-----掃描成功、PH平台沒空
                                Robot("0200", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                                PHQueue.Enqueue(Newly.Dequeue())
                            ElseIf PHMeasure.Count = 0 And PHPlatform.Status = "StandBy" Then '-----掃描成功、PH平台有空
                                Robot("0200", "0400")
                                Do Until SystemPLC.Read(3, 1) = "0001" Or SystemPLC.Read(3, 1) = "0002"
                                    Threading.Thread.Sleep(100)
                                Loop
                                Select Case SystemPLC.Read(3, 1)
                                    Case "0001"   '------------------------------------------掃描成功、PH平台有空、開蓋成功
                                        Robot("0400", "0502")
                                        PHMeasure.Enqueue(Newly.Dequeue())
                                        PHThread = New System.Threading.Thread(AddressOf PHFlow)
                                        PHThread.IsBackground = False
                                        PHThread.Start()
                                    Case "0002" '-------------------------------------------掃描成功、PH平台有空、開蓋失敗
                                        Robot("0400", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                                        Finish.Enqueue(Newly.Dequeue())
                                        BottleFinish(NewlyBottle)
                                        SeatStatus(Val(NewlyBottle.Name) - 100, Color.Red)
                                        MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & NewlyBottle.Name & "-OpenCan Failed")
                                        ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & NewlyBottle.Name & "-OpenCan Failed")
                                End Select
                            End If
                        Case "62"
                            Dim BottleFormula As JArray
                            BottleFormula = JArray.Parse(NewlyBottle.Formula)
                            Dim BottleResult = From objs In BottleFormula.Values(Of JObject)() Where objs("actVal").ToString() Select objs
                            If SGPlatform.Status = "ShutDown" Then '------------------------------------------------------掃描成功、SG平台關閉中
                                Robot("0200", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                                Finish.Enqueue(Newly.Dequeue())
                                BottleFinish(NewlyBottle)
                                SeatStatus(Val(NewlyBottle.Name) - 100, Color.Red)
                                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & NewlyBottle.Name & "-SGPlatform had been ShutDown")
                                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & NewlyBottle.Name & "-Platform had been ShutDown")
                            ElseIf BottleResult.Single.Item("actVal").ToString = 2 And SGPickling.Count <= SGQueue.Count Then '----------------------------掃描成功，沒有酸洗液可使用
                                Robot("0200", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                                Finish.Enqueue(Newly.Dequeue())
                                BottleFinish(NewlyBottle)
                                SeatStatus(Val(NewlyBottle.Name) - 100, Color.Red)
                                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & NewlyBottle.Name & "-No Enough PicklingBottle To Use")
                                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & NewlyBottle.Name & "-No Enough PicklingBottle To Use")
                            ElseIf SGMeasure.Count > 0 Then '------------------------------------------------------------掃描成功、SG平台沒空
                                Robot("0200", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                                SGQueue.Enqueue(Newly.Dequeue())
                            ElseIf SGMeasure.Count = 0 And SGPlatform.Status <> "StandBy" Then '-----掃描成功、SG平台沒空
                                Robot("0200", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                                SGQueue.Enqueue(Newly.Dequeue())
                            ElseIf SGMeasure.Count = 0 And SGPlatform.Status = "StandBy" Then '-----掃描成功、SG平台有空
                                Robot("0200", "0400")
                                Do Until SystemPLC.Read(3, 1) = "0001" Or SystemPLC.Read(3, 1) = "0002"
                                    Threading.Thread.Sleep(100)
                                Loop
                                Select Case SystemPLC.Read(3, 1)
                                    Case "0001"   '------------------------------------------掃描成功、SG平台有空、開蓋成功
                                        Robot("0400", "0501")
                                        SGMeasure.Enqueue(Newly.Dequeue())
                                        SGThread = New System.Threading.Thread(AddressOf SGFlow)
                                        SGThread.IsBackground = False
                                        SGThread.Start()
                                    Case "0002" '-------------------------------------------掃描成功、SG平台有空、開蓋失敗
                                        Robot("0400", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                                        Finish.Enqueue(Newly.Dequeue())
                                        BottleFinish(NewlyBottle)
                                        SeatStatus(Val(NewlyBottle.Name) - 100, Color.Red)
                                        MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & NewlyBottle.Name & "-OpenCan Failed")
                                        ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & NewlyBottle.Name & "-OpenCan Failed")
                                End Select
                            End If
                        Case "63"
                            If UVPlatform.Status = "ShutDown" Then '------------------------------------------------------掃描成功、UV平台關閉中
                                Robot("0200", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                                Finish.Enqueue(Newly.Dequeue())
                                BottleFinish(NewlyBottle)
                                SeatStatus(Val(NewlyBottle.Name) - 100, Color.Red)
                                MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & NewlyBottle.Name & "-UVPlatform had been ShutDown")
                                ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & NewlyBottle.Name & "-Platform had been ShutDown")
                            ElseIf UVMeasure.Count > 0 Then '------------------------------------------------------------掃描成功、UV平台沒空
                                Robot("0200", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                                UVQueue.Enqueue(Newly.Dequeue())
                            ElseIf UVMeasure.Count = 0 And UVPlatform.Status <> "StandBy" Then '-----掃描成功、UV平台沒空
                                Robot("0200", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                                UVQueue.Enqueue(Newly.Dequeue())
                            ElseIf UVMeasure.Count = 0 And UVPlatform.Status = "StandBy" Then '-----掃描成功、UV平台有空
                                Robot("0200", "0400")
                                Do Until SystemPLC.Read(3, 1) = "0001" Or SystemPLC.Read(3, 1) = "0002"
                                    Threading.Thread.Sleep(100)
                                Loop
                                Select Case SystemPLC.Read(3, 1)
                                    Case "0001"   '------------------------------------------掃描成功、UV平台有空、開蓋成功
                                        Robot("0400", "0503")
                                        UVMeasure.Enqueue(Newly.Dequeue())
                                        UVThread = New System.Threading.Thread(AddressOf UVFlow)
                                        UVThread.IsBackground = False
                                        UVThread.Start()
                                    Case "0002" '-------------------------------------------掃描成功、UV平台有空、開蓋失敗
                                        Robot("0400", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                                        Finish.Enqueue(Newly.Dequeue())
                                        BottleFinish(NewlyBottle)
                                        SeatStatus(Val(NewlyBottle.Name) - 100, Color.Red)
                                        MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & NewlyBottle.Name & "-OpenCan Failed")
                                        ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & NewlyBottle.Name & "-OpenCan Failed")
                                End Select
                            End If
                        Case Else
                            Robot("0200", NewlyBottle.Name.ToString.PadLeft(4, "0"))
                            Finish.Enqueue(Newly.Dequeue())
                            BottleFinish(NewlyBottle)
                            SeatStatus(Val(NewlyBottle.Name) - 100, Color.Red)
                            MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Bottle" & NewlyBottle.Name & "-No Data in Server : " & NewlyBottle.BottleId)
                            ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & NewlyBottle.Name & "-No Data in Server : " & NewlyBottle.BottleId)
                    End Select
            End Select
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("ScanNew Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub

    Sub TimerFirst(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ScanTimer.Elapsed
        Try
            ScanTimer.Stop()
            Scanner.Close()
            QRCodeStatus = 1
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("TimeFirst Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub ScanFirst() Handles Scanner.DataReceived
        Try
            Dim bottleCode As String = Scanner.ReadExisting
            Dim bottleResolution() As String = bottleCode.Split(",")
            QRCode = Mid(bottleResolution(2), 7, 9)
            ScanTimer.Stop()
            Scanner.Close()
            QRCodeStatus = 2
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("ScanFirst Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub

    Sub PHSoakFlow()
        Try
            '----------------------------------------------------------------------PHPLC
            PHSoakFlowPLC = New PLC
            PHSoakFlowPLC.CreatConnect(PLC_IP, PLC_Port)
            '----------------------------------------------------------------------PH站狀態設定
            PHPlatform.Status = "Domancy"
            '----------------------------------------------------------------------PH流程開始
            Threading.Thread.Sleep(5000)
            PHSoakFlowPLC.Write(200, "FFFF")
            Do Until PHQueue.Count > 0
                Threading.Thread.Sleep(100)
            Loop
            PHSoakFlowPLC.Write(200, "0000")
            PHPlatform.Status = "Remove"
            '----------------------------------------------------------------------
            Do Until PHPlatform.Status = "Washing"
                Threading.Thread.Sleep(100)
            Loop
            '----------------------------------------------------------------------
            PHSoakFlowPLC.Write(200, "0051")
            Do Until PHSoakFlowPLC.Read(200, 1) = "0052"
                Threading.Thread.Sleep(100)
            Loop
            '----------------------------------------------------------------------
            PHPlatform.Status = "StandBy"
            '----------------------------------------------------------------------PH物件釋放
            PHSoakFlowPLC.Dispose()
            PHSoakFlowPLC = Nothing
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("PHSoak Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub PHFlow()
        Try
            '----------------------------------------------------------------------PHPLC
            PHFlowPLC = New PLC
            PHFlowPLC.CreatConnect(PLC_IP, PLC_Port)
            '----------------------------------------------------------------------PH站狀態設定
            PHPlatform.Status = "Operationing"
            '----------------------------------------------------------------------PH站設定
            Dim PHDevice As SerialPort = New SerialPort
            Dim PHBottle As Bottle = PHMeasure.Peek
            PHBottle.AnalysisEndDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            '----------------------------------------------------------------------PH站流程開始
            PHFlowPLC.Write(200, "0001")
            '------------------------------
            Do Until PHFlowPLC.Read(200, 1) = "0004"
                Threading.Thread.Sleep(100)
            Loop
            Threading.Thread.Sleep(15000)
            PHDevice.PortName = "COM1"
            PHDevice.BaudRate = "9600"
            PHDevice.Parity = Parity.None
            PHDevice.DataBits = 8
            PHDevice.StopBits = StopBits.One
            If PHDevice.IsOpen = True Then
                PHDevice.Close()
            End If
            PHDevice.Open()
            Dim PHData As String = Nothing
            Dim Code As String = "010300350002"
            Code = CRC_Check(Code)
            Dim Data_str(7) As String
            Dim Data_Byte(7) As Byte
            For i = 0 To 7
                Data_str(i) = Mid(Code, i * 2 + 1, 2)
            Next
            For i = 0 To 7
                Data_Byte(i) = Convert.ToInt32(Data_str(i), 16)
            Next
            PHDevice.Write(Data_Byte, 0, 8)
            Threading.Thread.Sleep(300)
            Dim Data_Receive(8) As Byte
            PHDevice.Read(Data_Receive, 0, 9)
            Dim B(3)
            For i = 0 To 3
                B(i) = Hex(Data_Receive(i + 3)).PadLeft(2, "0")
            Next
            Dim data As String = HexToFloat(B(2) & B(3) & B(0) & B(1))
            PHBottle.Result = data
            PHDevice.Close()
            PHFlowPLC.Write(200, "0005")
            BottleLog(PHBottle.Json)
            '-------------------------
            Do Until PHFlowPLC.Read(200, 1) = "0006"
                Threading.Thread.Sleep(100)
            Loop
            PHMeasure.Dequeue()
            PHMeasure.Enqueue(PHBottle)
            PHPlatform.Status = "Remove"
            '-----------------------------
            Do Until PHPlatform.Status = "Washing"
                Threading.Thread.Sleep(100)
            Loop
            PHFlowPLC.Write(200, "0007")
            '----------------------------
            Do Until PHFlowPLC.Read(200, 1) = "0007"
                Threading.Thread.Sleep(100)
            Loop
            PHFlowPLC.Write(200, "0008")
            '-----------------------------
            Do Until PHFlowPLC.Read(200, 1) = "0009"
                Threading.Thread.Sleep(100)
            Loop
            PHPlatform.Status = "StandBy"
            '-------------------------------------------------------------PH物件釋放
            PHFlowPLC.Dispose()
            PHFlowPLC = Nothing
            PHDevice = Nothing
            PHBottle = Nothing
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("PH Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub UVFlow()
        Try
            '----------------------------------------------------------------------UVPLC
            UVFlowPLC = New PLC
            UVFlowPLC.CreatConnect(PLC_IP, PLC_Port)
            '----------------------------------------------------------------------UV站狀態設定
            UVPlatform.Status = "Operationing"
            '----------------------------------------------------------------------UV站設定
            Dim UV1280 As UV1280 = New UV1280
            UV1280.SetPort("COM5")
            Dim UVBottle As Bottle = UVMeasure.Peek
            UVBottle.AnalysisEndDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            '----------------------------------------------------------------------UV傳值
            Dim UVLength As String = Nothing
            Dim UVStr As String = UVBottle.Formula 'Nothing ' formula(WorkingBottleNo)
            Dim JArr As JArray = JArray.Parse(UVStr)
            For i = 0 To JArr.Count - 1
                Select Case JArr.ElementAt(i).Item("actSn").ToString
                    Case "1"    '----------樣本
                        Dim Value As String = Val(JArr.ElementAt(i).Item("actVal").ToString)
                        Dim int As Double
                        Dim fraction As Double
                        Dim asd As Integer = 5
                        Dim FFF(1) As Byte
                        If Value >= 5 Then
                            int = Math.Floor(Val(JArr.ElementAt(i).Item("actVal").ToString) / 5)
                            FFF = System.Text.Encoding.Default.GetBytes(int)
                            For k = 0 To 1
                                FFF(k) = Convert.ToString(FFF(k), 16)
                            Next
                            UVFlowPLC.Write(312, FFF(0) & FFF(1))
                        ElseIf Value < 5 Then
                            int = 0
                            UVFlowPLC.Write(312, int.ToString.PadLeft(4, "0"))
                        End If
                        fraction = Val(Value) Mod asd
                        Dim pace As String = (fraction * 1200).ToString.PadLeft(4, "0")
                        Dim pace_Str() As Byte = System.Text.Encoding.Default.GetBytes(pace)
                        Dim pace_Hex(3) As String
                        For j = 0 To 3
                            pace_Hex(j) = Convert.ToString(pace_Str(j), 16)
                        Next
                        UVFlowPLC.Write(313, pace_Hex(0) & pace_Hex(1))
                        UVFlowPLC.Write(314, pace_Hex(2) & pace_Hex(3))
                    Case "2"    '----------純水
                        Dim Value As String = Val(JArr.ElementAt(i).Item("actVal").ToString)
                        Dim int As String
                        Dim fraction As Double
                        Dim asd As Integer = 5
                        Dim FFF(1) As Byte
                        If Value >= 5 Then
                            int = Math.Floor(Val(JArr.ElementAt(i).Item("actVal").ToString) / 5).ToString.PadLeft(2, "0")
                            FFF = System.Text.Encoding.Default.GetBytes(int)
                            For k = 0 To 1
                                FFF(k) = Convert.ToString(FFF(k), 16)
                            Next
                            UVFlowPLC.Write(322, FFF(0) & FFF(1))
                        ElseIf Value < 5 Then
                            int = 0
                            UVFlowPLC.Write(322, int.ToString.PadLeft(4, "0"))
                        End If
                        fraction = Val(Value) Mod asd
                        Dim pace As String = (fraction * 1200).ToString.PadLeft(4, "0")
                        Dim pace_Str() As Byte = System.Text.Encoding.Default.GetBytes(pace)
                        Dim pace_Hex(3) As String
                        For j = 0 To 3
                            pace_Hex(j) = Convert.ToString(pace_Str(j), 16)
                        Next
                        UVFlowPLC.Write(323, pace_Hex(0) & pace_Hex(1))
                        UVFlowPLC.Write(324, pace_Hex(2) & pace_Hex(3))
                    Case "65"   '----------KOH
                        Dim Value As String = Val(JArr.ElementAt(i).Item("actVal").ToString)
                        Dim int As String
                        Dim fraction As Double
                        Dim asd As Integer = 5
                        Dim FFF(1) As Byte
                        If Value >= 5 Then
                            int = Math.Floor(Val(JArr.ElementAt(i).Item("actVal").ToString) / 5).ToString.PadLeft(2, "0")
                            FFF = System.Text.Encoding.Default.GetBytes(int)
                            For k = 0 To 1
                                FFF(k) = Convert.ToString(FFF(k), 16)
                            Next
                            UVFlowPLC.Write(332, FFF(0) & FFF(1))
                        ElseIf Value < 5 Then
                            int = 0
                            UVFlowPLC.Write(332, int.ToString.PadLeft(4, "0"))
                        End If
                        fraction = Val(Value) Mod asd
                        Dim pace As String = (fraction * 1200).ToString.PadLeft(4, "0")
                        Dim pace_Str() As Byte = System.Text.Encoding.Default.GetBytes(pace)
                        Dim pace_Hex(3) As String
                        For j = 0 To 3
                            pace_Hex(j) = Convert.ToString(pace_Str(j), 16)
                        Next
                        UVFlowPLC.Write(333, pace_Hex(0) & pace_Hex(1))
                        UVFlowPLC.Write(334, pace_Hex(2) & pace_Hex(3))
                    Case "64"   '-----------確認緩衝液針閥
                        Dim Value As String = Val(JArr.ElementAt(i).Item("prepareId_buffer").ToString)
                        If Value = "1" Then
                            UVFlowPLC.Write(311, "0001")
                        ElseIf Value = "2" Then
                            UVFlowPLC.Write(321, "0001")
                        ElseIf Value = "65" Then
                            UVFlowPLC.Write(331, "0001")
                        End If
                    Case "63"
                        UVLength = JArr.ElementAt(i).Item("actVal").ToString
                        UVBottle.UVLength = UVLength
                    Case Else
                End Select
            Next
            '----------------------------------------------------------------------UV站流程開始
            UVFlowPLC.Write(300, "0001")
            '-------------------------------------------------
            Do Until UVFlowPLC.Read(300, 1) = "0004"
                Threading.Thread.Sleep(100)
            Loop
            For i = 0 To 15
                UV1280.Motor()
            Next
            UVFlowPLC.Write(300, "0005")
            '--------------------------------------------------
            Do Until UVFlowPLC.Read(300, 1) = "0006"
                Threading.Thread.Sleep(100)
            Loop
            For i = 0 To 5
                UV1280.Motor()
            Next
            UVFlowPLC.Write(300, "0007")
            '---------------------------------------------------
            Do Until UVFlowPLC.Read(300, 1) = "0008"
                Threading.Thread.Sleep(100)
            Loop
            For i = 0 To 2
                UV1280.Motor()
            Next
            UVFlowPLC.Write(300, "0009")
            '------------------------------------------------------
            Do Until UVFlowPLC.Read(300, 1) = "0010"
                Threading.Thread.Sleep(100)
            Loop
            UV1280.SetWaveLength(Val(UVLength) * 10)
            UV1280.Reset()
            UVBottle.Blank = UV1280.GetAbsorb
            UVFlowPLC.Write(300, "0011")
            '-------------------------------------------------------
            Do Until UVFlowPLC.Read(300, 1) = "0012"
                Threading.Thread.Sleep(100)
            Loop
            For i = 0 To 5
                UV1280.Motor()
            Next
            UVFlowPLC.Write(300, "0013")
            '--------------------------------------------------------
            Do Until UVFlowPLC.Read(300, 1) = "0015"
                Threading.Thread.Sleep(100)
            Loop
            For i = 0 To 6
                UV1280.Motor()
            Next
            UVFlowPLC.Write(300, "0016")
            '-----------------------------------------------------------
            Do Until UVFlowPLC.Read(300, 1) = "0017"
                Threading.Thread.Sleep(100)
            Loop
            UVBottle.Result = UV1280.GetAbsorb
            UVFlowPLC.Write(300, "0018")
            BottleLog(UVBottle.Json)
            '---------------------------------------------------------
            Do Until UVFlowPLC.Read(300, 1) = "0019"
                Threading.Thread.Sleep(100)
            Loop
            UVMeasure.Dequeue()
            UVMeasure.Enqueue(UVBottle)
            UVPlatform.Status = "Remove"
            '---------------------------------------------------------
            Do Until UVPlatform.Status = "Washing"
                Threading.Thread.Sleep(100)
            Loop
            UVFlowPLC.Write(300, "0020")
            '-------------------------------------------------------
            Do Until UVFlowPLC.Read(300, 1) = "0022"
                Threading.Thread.Sleep(100)
            Loop
            For i = 0 To 5
                UV1280.Motor()
            Next
            UVFlowPLC.Write(300, "0023")
            '--------------------------------------------------
            Do Until UVFlowPLC.Read(300, 1) = "0024"
                Threading.Thread.Sleep(100)
            Loop
            For i = 0 To 20
                UV1280.Motor()
            Next
            UVFlowPLC.Write(300, "0025")
            UVPlatform.Status = "StandBy"
            '----------------------------------------------------------------------UV物件釋放
            UVFlowPLC.Dispose()
            UVFlowPLC = Nothing
            UV1280 = Nothing
            UVBottle = Nothing
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("UV Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub SGFlow()
        Try '
            '----------------------------------------------------------------------SGPLC
            SGFlowPLC = New PLC
            SGFlowPLC.CreatConnect(PLC_IP, PLC_Port)
            '----------------------------------------------------------------------SG站狀態設定
            SGPlatform.Status = "Operationing"
            '----------------------------------------------------------------------SG站設定
            Dim SGDevice As SerialPort = New SerialPort
            Dim SGBottle As Bottle = SGMeasure.Peek
            SGBottle.AnalysisEndDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Dim BottleFormula As JArray
            BottleFormula = JArray.Parse(SGBottle.Formula)
            Dim BottleResult = From objs In BottleFormula.Values(Of JObject)() Where objs("actVal").ToString() Select objs
            '----------------------------------------------------------------------SG站流程開始
            SGFlowPLC.Write(100, "0001")
            '---------------------------------------
            Do Until SGFlowPLC.Read(100, 1) = "0005"
                Threading.Thread.Sleep(100)
            Loop
            SGDevice.PortName = "COM3"
            SGDevice.BaudRate = 9600
            SGDevice.Parity = Parity.None
            SGDevice.DataBits = 8
            SGDevice.StopBits = StopBits.One
            If SGDevice.IsOpen = True Then
                SGDevice.Close()
            End If
            SGDevice.Open()
            Dim SGData As String = Nothing
            Dim SGValue() As String = Nothing
            SGData = SGDevice.ReadLine
            Threading.Thread.Sleep(100)
            SGValue = SGData.Split("=")
            Do Until SGValue.Count = 3
                SGData = SGDevice.ReadLine
                SGValue = SGData.Split("=")
            Loop
            SGBottle.Result = SGValue(2).Substring(1, 6)
            SGDevice.Close()
            SGFlowPLC.Write(100, "0006")
            BottleLog(SGBottle.Json)
            '----------------------------------------
            Do Until SGFlowPLC.Read(100, 1) = "0007"
                Threading.Thread.Sleep(100)
            Loop
            SGMeasure.Dequeue()
            SGMeasure.Enqueue(SGBottle)
            SGPlatform.Status = "Remove"
            '-----------------------------------------
            Do Until SGPlatform.Status = "Washing"
                Threading.Thread.Sleep(100)
            Loop
            SGFlowPLC.Write(100, "0008")
            '-----------------------------------------
            Do Until SGFlowPLC.Read(100, 1) = "0008"
                Threading.Thread.Sleep(100)
            Loop
            SGFlowPLC.Write(100, "0009")
            '-----------------------------------------
            Do Until SGFlowPLC.Read(100, 1) = "0010"
                Threading.Thread.Sleep(100)
            Loop
            SGFlowPLC.Write(100, "0011")
            Select Case BottleResult.Single.Item("actVal").ToString
                Case "1"
                    SGPlatform.Status = "StandBy"
                Case "2"
                    SGPlatform.Status = "Pickling"
            End Select
            '---------------------------------------------------------------SG物件釋放
            SGFlowPLC.Dispose()
            SGFlowPLC = Nothing
            SGDevice = Nothing
            SGBottle = Nothing
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("SG Function")
            MsgBox("Something Error, Please restart the system.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub SGPicklingFlow()
        Try
            '----------------------------------------------------------------------SGPLC
            SGPicklingFlowPLC = New PLC
            SGPicklingFlowPLC.CreatConnect(PLC_IP, PLC_Port)
            '----------------------------------------------------------------------SG站狀態設定
            SGPlatform.Status = "Operationing"
            '----------------------------------------------------------------------SG站流程開始
            SGPicklingFlowPLC.Write(100, "0051")
            '---------------------------------------
            Do Until SGPicklingFlowPLC.Read(100, 1) = "0055"
                Threading.Thread.Sleep(100)
            Loop
            SGPlatform.Status = "Remove"
            '-----------------------------------------
            Do Until SGPlatform.Status = "Washing"
                Threading.Thread.Sleep(100)
            Loop
            SGPicklingFlowPLC.Write(100, "0056")
            '-----------------------------------------
            Do Until SGPicklingFlowPLC.Read(100, 1) = "0056"
                Threading.Thread.Sleep(100)
            Loop
            SGPicklingFlowPLC.Write(100, "0057")
            '-----------------------------------------
            Do Until SGPicklingFlowPLC.Read(100, 1) = "0058"
                Threading.Thread.Sleep(100)
            Loop
            SGPicklingFlowPLC.Write(100, "0059")
            SGPlatform.Status = "StandBy"
            DataBase.UpdateVolume("19-43", "70")
            '---------------------------------------------------------------SG物件釋放
            SGPicklingFlowPLC.Dispose()
            SGPicklingFlowPLC = Nothing
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("SGPickling Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub

  

   
    Sub CheckSeat() '------------------------該盤全部結束後移出Queue，Button.text => Finish
        Try
            For i = 0 To Seat.Count - 1
                Select Case Seat.Peek
                    Case 1
                        If Seat1.BottleCount = Seat1.BottleFinish Then
                            For j = 0 To Seat1.BottleCount - 1
                                Dim Bottle1 As Bottle = Seat1.Bottles(j)
                                Dim Bottle2 As Bottle = Finish.Peek
                                Do Until Bottle2.Name = Bottle1.Name
                                    Finish.Enqueue(Finish.Dequeue)
                                    Bottle2 = Finish.Peek
                                Loop
                                Finish.Dequeue()
                            Next
                            Seat.Dequeue()
                            BTNSEAT1.Text = "Seat1 Finish"
                            BTNSEAT1.ForeColor = Color.Red
                            MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Seat1 Finish。")
                            ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "Seat1 Finish。")
                            Seat1 = Nothing
                        Else
                            Seat.Enqueue(Seat.Dequeue)
                        End If
                    Case 2
                        If Seat2.BottleCount = Seat2.BottleFinish Then
                            For j = 0 To Seat2.BottleCount - 1
                                Dim Bottle1 As Bottle = Seat2.Bottles(j)
                                Dim Bottle2 As Bottle = Finish.Peek
                                Do Until Bottle2.Name = Bottle1.Name
                                    Finish.Enqueue(Finish.Dequeue)
                                    Bottle2 = Finish.Peek
                                Loop
                                Finish.Dequeue()
                            Next
                            Seat.Dequeue()
                            BTNSEAT2.Text = "Seat2 Finish"
                            BTNSEAT2.ForeColor = Color.Red
                            MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Seat2 Finish。")
                            ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "Seat2 Finish。")
                            Seat2 = Nothing
                        Else
                            Seat.Enqueue(Seat.Dequeue)
                        End If
                    Case 3
                        If Seat3.BottleCount = Seat3.BottleFinish Then
                            For j = 0 To Seat3.BottleCount - 1
                                Dim Bottle1 As Bottle = Seat3.Bottles(j)
                                Dim Bottle2 As Bottle = Finish.Peek
                                Do Until Bottle2.Name = Bottle1.Name
                                    Finish.Enqueue(Finish.Dequeue)
                                    Bottle2 = Finish.Peek
                                Loop
                                Finish.Dequeue()
                            Next
                            Seat.Dequeue()
                            BTNSEAT3.Text = "Seat3 Finish"
                            BTNSEAT3.ForeColor = Color.Red
                            MessageDataTable.Rows.Add(Date.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Seat3 Finish。")
                            ActionRecord(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & "   " & "Seat3 Finish。")
                            Seat3 = Nothing
                        Else
                            Seat.Enqueue(Seat.Dequeue)
                        End If
                End Select
            Next
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("CheckSeat Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub SeatStatus()
        Try
            SeatPLC = New PLC
            SeatPLC.CreatConnect(PLC_IP, PLC_Port)
            Do
                Dim Seat1Hex As String = SeatPLC.Read(20001, 1)
                Dim Seat1DEC As Integer = Convert.ToInt32(Seat1Hex, 16)
                Dim Seat1B As String = Convert.ToString(Seat1DEC, 2).PadLeft(12, "0")
                Dim Seat1(Seat1B.Length - 1) As String
                For i = 0 To Seat1.Count - 1
                    Seat1(i) = Mid(Seat1B, i + 1, 1)
                Next
                Array.Reverse(Seat1)
                For i = 0 To Seat1.Count - 1
                    If Seat1(i) = 1 Then
                        CType(Me.TableLayoutPanel2.Controls("SeatStatus" & i + 1), Label).BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
                    Else
                        CType(Me.TableLayoutPanel2.Controls("SeatStatus" & i + 1), Label).BackColor = System.Drawing.Color.FromArgb(255, 240, 240, 240)
                    End If
                Next
                Dim Seat2Hex As String = SeatPLC.Read(20002, 1)
                Dim Seat2DEC As Integer = Convert.ToInt32(Seat2Hex, 16)
                Dim Seat2B As String = Convert.ToString(Seat2DEC, 2).PadLeft(12, "0")
                Dim Seat2(Seat2B.Length - 1) As String
                For i = 0 To Seat2.Count - 1
                    Seat2(i) = Mid(Seat2B, i + 1, 1)
                Next
                Array.Reverse(Seat2)
                For i = 0 To Seat2.Count - 1
                    If Seat2(i) = 1 Then
                        CType(Me.TableLayoutPanel4.Controls("SeatStatus" & i + 13), Label).BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
                    Else
                        CType(Me.TableLayoutPanel4.Controls("SeatStatus" & i + 13), Label).BackColor = System.Drawing.Color.FromArgb(255, 240, 240, 240)
                    End If
                Next
                Dim Seat3Hex As String = SeatPLC.Read(20003, 1)
                Dim Seat3DEC As Integer = Convert.ToInt32(Seat3Hex, 16)
                Dim Seat3B As String = Convert.ToString(Seat3DEC, 2).PadLeft(12, "0")
                Dim Seat3(Seat3B.Length - 1) As String
                For i = 0 To Seat3.Count - 1
                    Seat3(i) = Mid(Seat3B, i + 1, 1)
                Next
                Array.Reverse(Seat3)
                For i = 0 To Seat3.Count - 1
                    If Seat3(i) = 1 Then
                        CType(Me.TableLayoutPanel1.Controls("SeatStatus" & i + 25), Label).BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
                    Else
                        CType(Me.TableLayoutPanel1.Controls("SeatStatus" & i + 25), Label).BackColor = System.Drawing.Color.FromArgb(255, 240, 240, 240)
                    End If
                Next
                Dim Buffer1Hex As String = SeatPLC.Read(20004, 1)
                Dim Buffer1DEC As Integer = Convert.ToInt32(Buffer1Hex, 16)
                Dim Buffer1B As String = Convert.ToString(Buffer1DEC, 2).PadLeft(12, "0")
                Dim Buffer1(Buffer1B.Length - 1) As String
                For i = 0 To Buffer1.Count - 1
                    Buffer1(i) = Mid(Buffer1B, i + 1, 1)
                Next
                Array.Reverse(Buffer1)
                For i = 0 To Buffer1.Count - 1
                    If Buffer1(i) = 1 Then
                        CType(Me.TableLayoutPanel5.Controls("BufferStatus" & i + 1), Label).BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
                    Else
                        CType(Me.TableLayoutPanel5.Controls("BufferStatus" & i + 1), Label).BackColor = System.Drawing.Color.FromArgb(255, 240, 240, 240)
                    End If
                Next
                Threading.Thread.Sleep(50)
            Loop
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("Seat Status Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub ShowView()
        Try
            ListBox1.Items.Clear()
            ListBox2.Items.Clear()
            ListBox3.Items.Clear()
            ListBox4.Items.Clear()
            ListBox5.Items.Clear()
            ListBox6.Items.Clear()
            ListBox7.Items.Clear()
            ListBox8.Items.Clear()
            ListBox11.Items.Clear()
            ListBox12.Items.Clear()
            Dim NewlyShow(Newly.Count - 1) As Bottle
            Dim UVMeasureShow(UVMeasure.Count - 1) As Bottle
            Dim UVQueueShow(UVQueue.Count - 1) As Bottle
            Dim SGMeasureShow(SGMeasure.Count - 1) As Bottle
            Dim SGQueueShow(SGQueue.Count - 1) As Bottle
            Dim PHMeasureShow(PHMeasure.Count - 1) As Bottle
            Dim PHQueueShow(PHQueue.Count - 1) As Bottle
            Dim FinishShow(Finish.Count - 1) As Bottle
            Dim SGPicklingShow(SGPickling.Count - 1) As Bottle
            Dim PHSoakShow(PHSoak.Count - 1) As Bottle
            Newly.CopyTo(NewlyShow, 0)
            UVMeasure.CopyTo(UVMeasureShow, 0)
            UVQueue.CopyTo(UVQueueShow, 0)
            SGMeasure.CopyTo(SGMeasureShow, 0)
            SGQueue.CopyTo(SGQueueShow, 0)
            PHMeasure.CopyTo(PHMeasureShow, 0)
            PHQueue.CopyTo(PHQueueShow, 0)
            Finish.CopyTo(FinishShow, 0)
            SGPickling.CopyTo(SGPicklingShow, 0)
            PHSoak.CopyTo(PHSoakShow, 0)
            For i = 0 To NewlyShow.Count - 1
                ListBox1.Items.Add(NewlyShow(i).Name)
            Next
            For i = 0 To UVMeasureShow.Count - 1
                ListBox2.Items.Add(UVMeasureShow(i).Name)
            Next
            For i = 0 To UVQueue.Count - 1
                ListBox3.Items.Add(UVQueueShow(i).Name)
            Next
            For i = 0 To SGMeasureShow.Count - 1
                ListBox6.Items.Add(SGMeasureShow(i).Name)
            Next
            For i = 0 To SGQueueShow.Count - 1
                ListBox7.Items.Add(SGQueue(i).name)
            Next
            For i = 0 To PHMeasureShow.Count - 1
                ListBox4.Items.Add(PHMeasureShow(i).Name)
            Next
            For i = 0 To PHQueueShow.Count - 1
                ListBox5.Items.Add(PHQueueShow(i).Name)
            Next
            For i = 0 To FinishShow.Count - 1
                ListBox8.Items.Add(FinishShow(i).Name)
            Next
            For i = 0 To SGPicklingShow.Count - 1
                ListBox11.Items.Add(SGPickling(i).Name)
            Next
            For i = 0 To PHSoakShow.Count - 1
                ListBox12.Items.Add(PHSoakShow(i).Name)
            Next
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("ShowView Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub

    Sub BottleFinish(ByVal Bottle As Bottle)
        Try
            If 101 <= (Bottle.Name) And (Bottle.Name) <= 112 Then
                Seat1.BottleFinish = Seat1.BottleFinish + 1
            ElseIf 113 <= Val(Bottle.Name) And (Bottle.Name) <= 124 Then
                Seat2.BottleFinish = Seat2.BottleFinish + 1
            ElseIf 125 <= Val(Bottle.Name) And (Bottle.Name) <= 136 Then
                Seat3.BottleFinish = Seat3.BottleFinish + 1
            End If
        Catch ex As Exception
            ErrorLog(ex.Message.ToString)
            ErrorLog("Bottle Finish Function")
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
        
    End Sub
    Sub Robot(ByVal Start As String, ByVal Target As String)
        Try
            RobotPLC = New PLC
            RobotPLC.CreatConnect(PLC_IP, PLC_Port)
            RobotPLC.Write(0, Start)
            RobotPLC.Write(1, Target)
            Do Until RobotPLC.Read(2, 1) = Target
                Threading.Thread.Sleep(100)
            Loop
            RobotPLC.Write(0, "0000")
            RobotPLC.Write(1, "0000")
            RobotPLC.Dispose()
            RobotPLC = Nothing
        Catch ex As Exception
            ErrorLog("Robot Function")
            ErrorLog(ex.Message.ToString)
            MsgBox("Something Error, The system will be restart.")
            Dim a() = Process.GetProcessesByName("AutoLab1_2")
            If a.Count > 0 Then
                For i = 0 To a.Count - 1
                    a(i).Kill()
                Next
            End If
        End Try
    End Sub
    Sub SeatStatus(ByVal i As Integer, ByVal Color As Color)
        If 1 <= i And i <= 12 Then
            CType(Me.TableLayoutPanel2.Controls("SeatStatus" & i), Label).ForeColor = Color
        ElseIf 13 <= i And i <= 24 Then
            CType(Me.TableLayoutPanel4.Controls("SeatStatus" & i), Label).ForeColor = Color
        ElseIf 25 <= i And i <= 36 Then
            CType(Me.TableLayoutPanel1.Controls("SeatStatus" & i), Label).ForeColor = Color
        End If
    End Sub
    Sub BottleLog(ByVal SendData As String)
        Dim Result As String = DataBase.sendData(SendData)
        Dim File As System.IO.StreamWriter
        File = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\Log\" & Date.Now.ToString("yyyyMMdd") & "_BottleLog.txt", True)
        File.WriteLine(SendData & vbCrLf & Result)
        File.Close()
    End Sub
    Sub ActionRecord(ByVal Action As String)
        Dim File As System.IO.StreamWriter
        File = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\Log\" & Date.Now.ToString("yyyyMMdd") & "_ActionLog.txt", True)
        File.WriteLine(Action)
        File.Close()
    End Sub
    Sub ErrorLog(ByVal Name As String)
        Dim File As System.IO.StreamWriter
        File = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\Log\" & Date.Now.ToString("yyyyMMdd") & "_ErrorLog.txt", True)
        File.WriteLine(Name)
        File.Close()
    End Sub
    Function HexToFloat(ByVal Hex As String)
        Dim Finish As String
        Dim InputIndex As Integer = 0
        Dim OutputIndex As Integer = 0
        Dim bArray(3) As Byte
        Dim HexValue = Hex
        For InputIndex = 0 To HexValue.Length - 1 Step 2
            bArray(OutputIndex) = Byte.Parse(HexValue.Chars(InputIndex) & HexValue.Chars(InputIndex + 1), Globalization.NumberStyles.HexNumber)
            OutputIndex += 1
        Next
        Array.Reverse(bArray)
        Finish = BitConverter.ToSingle(bArray, 0).ToString()
        Return Finish
    End Function
    Function CRC_Check(ByVal CRCString As String) As String
        Dim CRC16(1) As Byte
        Dim CRC(1) As Byte
        Dim mOutData(Len(CRCString) / 2 - 1) As Byte
        Dim X0, X1 As Integer
        For i = 0 To Len(CRCString) / 2 - 1
            mOutData(i) = Val("&H" & Mid(CRCString, i * 2 + 1, 2))
        Next i
        CRC16(0) = 1
        CRC16(1) = 160
        CRC(0) = 255
        CRC(1) = 255
        For i = 0 To UBound(mOutData)
            CRC(0) = CRC(0) Xor mOutData(i)
            For J = 1 To 8
                X1 = CRC(1) Mod 2
                CRC(1) = Int(CRC(1) / 2)
                X0 = CRC(0) Mod 2
                CRC(0) = Int(CRC(0) / 2)
                If X1 = 1 Then
                    CRC(0) = CRC(0) + 128
                End If
                If X0 = 1 Then
                    CRC(0) = CRC(0) Xor CRC16(0)
                    CRC(1) = CRC(1) Xor CRC16(1)
                End If
            Next
        Next
        If Len(CStr(Hex(CRC(0)))) = 1 Then
            CRC_Check = CRCString + "0" + Hex(CRC(0))
        Else
            CRC_Check = CRCString + Hex(CRC(0))
        End If
        If Len(CStr(Hex(CRC(1)))) = 1 Then
            CRC_Check = CRC_Check + "0" + Hex(CRC(1))
        Else
            CRC_Check = CRC_Check + Hex(CRC(1))
        End If
    End Function
End Class
