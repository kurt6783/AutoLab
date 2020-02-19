Imports Newtonsoft.Json.Linq
Imports System.Net.NetworkInformation
Imports System.Net
Imports System.IO
Imports System.Text
Public Class DataBase
    Dim DownLoadPath As String = "http://192.168.50.102/system_mvc/controller.php?s=modules_mvc,F50,laboratoryRecord,mobile_20190605&action=getPrepare_ac&unVal=1"
    Dim UpLoadPath As String = "http://192.168.50.102/system_mvc/controller.php?s=modules_mvc,F50,laboratoryRecord,mobile&action=updateVal"
    Dim VolumePath As String = "http://192.168.50.102/system_mvc/controller.php?s=modules_mvc,F50,laboratoryRecord,mobile&action=portUse"
    Dim myPS As String = "{""mac"":""" & getMacAddress(0) & """}"
    Dim BottleStr As String
    Dim BottleData As JArray
    Dim DetectionStr As String
    Dim DetectionData As JArray
    Dim FormulaStr As String
    Dim FormulaData As JArray
    Sub ConnectDB()
        BottleStr = GetData(DownLoadPath, myPS)
        BottleData = JArray.Parse(BottleStr)
    End Sub
    Function tankSn(ByVal bottleID As String)
        Try
            Dim BottleResult = From objs In BottleData.Values(Of JObject)() Where objs("bottleId").ToString() = bottleID Select objs
            Return BottleResult.Single.Item("tankSn").ToString
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Function setDate(ByVal bottleID As String)
        Try
            Dim BottleResult = From objs In BottleData.Values(Of JObject)() Where objs("bottleId").ToString() = bottleID Select objs
            Return BottleResult.Single.Item("setDate").ToString
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Function itemSn(ByVal BottleID As String)
        Try
            Dim BottleResult = From objs In BottleData.Values(Of JObject)() Where objs("bottleId").ToString() = BottleID Select objs
            Dim Items_Str As String = BottleResult.Single.Item("items").ToString
            Dim ItemsData As JArray = JArray.Parse(Items_Str)
            Return ItemsData.Single.Item("itemSn").ToString
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Function onmicSn(ByVal BottleID As String)
        Try
            Dim BottleResult = From objs In BottleData.Values(Of JObject)() Where objs("bottleId").ToString() = BottleID Select objs
            Dim Items_Str As String = BottleResult.Single.Item("items").ToString
            Dim ItemsData As JArray = JArray.Parse(Items_Str)
            Return ItemsData.Single.Item("onmicSn").ToString
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Function methodId(ByVal BottleID As String)
        Try
            Dim BottleResult = From objs In BottleData.Values(Of JObject)() Where objs("bottleId").ToString() = BottleID Select objs
            Dim Items_Str As String = BottleResult.Single.Item("items").ToString
            Dim ItemsData As JArray = JArray.Parse(Items_Str)
            Return ItemsData.Single.Item("methodId").ToString
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Function methodSn(ByVal BottleID As String)
        Try
            Dim BottleResult = From objs In BottleData.Values(Of JObject)() Where objs("bottleId").ToString() = BottleID Select objs
            Dim Items_Str As String = BottleResult.Single.Item("items").ToString
            Dim ItemsData As JArray = JArray.Parse(Items_Str)
            Return ItemsData.Single.Item("methodSn").ToString
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Function formula(ByVal BottleID As String)
        Try
            Dim BottleResult = From objs In BottleData.Values(Of JObject)() Where objs("bottleId").ToString() = BottleID Select objs
            Dim Items_Str As String = BottleResult.Single.Item("items").ToString
            Dim ItemsData As JArray = JArray.Parse(Items_Str)
            Return ItemsData.Single.Item("method").ToString
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Function sendData(ByVal myString As String)
        Dim request As WebRequest = WebRequest.Create(UpLoadPath)
        request.Method = "POST"
        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(myString)
        request.ContentType = "application/x-www-form-urlencoded"
        request.ContentLength = byteArray.Length
        Dim dataStream As Stream = request.GetRequestStream()
        dataStream.Write(byteArray, 0, byteArray.Length)
        dataStream.Close()
        Dim response As WebResponse = request.GetResponse()
        Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
        dataStream = response.GetResponseStream()
        Dim reader As New StreamReader(dataStream)
        Dim responseFromServer As String = reader.ReadToEnd()
        Console.WriteLine(responseFromServer)
        reader.Close()
        dataStream.Close()
        response.Close()
        Return responseFromServer
    End Function
    Private Function GetData(ByVal URLPath As String, ByVal myString As String)
        Dim request As WebRequest = WebRequest.Create(URLPath)
        request.Method = "POST"
        Dim postData As String = ""
        postData = "json=" & myString
        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)
        request.ContentType = "application/x-www-form-urlencoded"
        request.ContentLength = byteArray.Length
        request.Timeout = 10000
        Dim dataStream As Stream = request.GetRequestStream()
        dataStream.Write(byteArray, 0, byteArray.Length)
        dataStream.Close()
        Dim response As WebResponse = request.GetResponse()
        Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
        dataStream = response.GetResponseStream()
        Dim reader As New StreamReader(dataStream)
        Dim responseFromServer As String = reader.ReadToEnd()
        Console.WriteLine(responseFromServer)
        reader.Close()
        dataStream.Close()
        response.Close()
        Return responseFromServer
    End Function
    Function UpdateVolume(ByVal Port As String, ByVal volume As String)
        Dim myString As String
        Dim Obj As New JObject
        Dim Arr As New JArray
        Obj.Add(Port, volume)
        Arr.Add(Obj)
        myString = Arr.ToString
        myString = "json=" & myString
        Dim request As WebRequest = WebRequest.Create(VolumePath)
        request.Method = "POST"
        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(myString)
        request.ContentType = "application/x-www-form-urlencoded"
        request.ContentLength = byteArray.Length
        Dim dataStream As Stream = request.GetRequestStream()
        dataStream.Write(byteArray, 0, byteArray.Length)
        dataStream.Close()
        Dim response As WebResponse = request.GetResponse()
        Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
        dataStream = response.GetResponseStream()
        Dim reader As New StreamReader(dataStream)
        Dim responseFromServer As String = reader.ReadToEnd()
        Console.WriteLine(responseFromServer)
        reader.Close()
        dataStream.Close()
        response.Close()
        Return responseFromServer
    End Function
    Private Function getMacAddress(ByVal index As Integer)
        Dim nics() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces
        Dim tempStr As String = nics(index).GetPhysicalAddress.ToString
        Return tempStr(0) & tempStr(1) & ":" & tempStr(2) & tempStr(3) & ":" & tempStr(4) & tempStr(5) & ":" & tempStr(6) & tempStr(7) & ":" & tempStr(8) & tempStr(9) & ":" & tempStr(10) & tempStr(11)
    End Function
End Class
