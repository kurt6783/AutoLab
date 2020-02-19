Imports Newtonsoft.Json.Linq

Public Class Bottle
    Private myStatus As String
    Private myType As String
    Private myName As String
    Private myBottleId As String
    Private myTankSn As String
    Private mySetDate As String
    Private myItemSn As String
    Private myOnmicSn As String
    Private myItem As String
    Private myMethodId As String
    Private myMethodSn As String
    Private myFormula As String
    Private myMac As String
    Private myPreprocessDate As String
    Private myAnalysisEndDate As String
    Private myResult As String
    Private myJson As String
    Private myUVLength As String
    Private myBlank As String
    Public Property Type()
        Get
            Return myType
        End Get
        Set(ByVal value)
            myType = value
        End Set
    End Property
    Public Property Name()
        Get
            Return myName
        End Get
        Set(ByVal value)
            myName = value
        End Set
    End Property
    Public Property BottleId()
        Get
            Return myBottleId
        End Get
        Set(ByVal value)
            myBottleId = value
        End Set
    End Property
    Public Property TankSn()
        Get
            Return myTankSn
        End Get
        Set(ByVal value)
            myTankSn = value
        End Set
    End Property
    Public Property SetDate()
        Get
            Return mySetDate
        End Get
        Set(ByVal value)
            mySetDate = value
        End Set
    End Property
    Public Property ItemSn()
        Get
            Return myItemSn
        End Get
        Set(ByVal value)
            myItemSn = value
        End Set
    End Property
    Public Property OnmicSn()
        Get
            Return myOnmicSn
        End Get
        Set(ByVal value)
            myOnmicSn = value
        End Set
    End Property
    Public Property Item()
        Get
            Return myItem
        End Get
        Set(ByVal value)
            myItem = value
        End Set
    End Property
    Public Property MethodId()
        Get
            Return myMethodId
        End Get
        Set(ByVal value)
            myMethodId = value
        End Set
    End Property
    Public Property MethodSn()
        Get
            Return myMethodSn
        End Get
        Set(ByVal value)
            myMethodSn = value
        End Set
    End Property
    Public Property Formula() As String
        Get
            Return myFormula
        End Get
        Set(ByVal value As String)
            myFormula = value
        End Set
    End Property
    Public Property Mac()
        Get
            myMac = "74:D4:35:BB:D7:24"
            Return myMac
        End Get
        Set(ByVal value)
            myMac = value
        End Set
    End Property
    Public Property PreprocessDate()
        Get
            Return myPreprocessDate
        End Get
        Set(ByVal value)
            myPreprocessDate = value
        End Set
    End Property
    Public Property AnalysisEndDate()
        Get
            Return myAnalysisEndDate
        End Get
        Set(ByVal value)
            myAnalysisEndDate = value
        End Set
    End Property
    Public Property Result()
        Get
            Return myResult
        End Get
        Set(ByVal value)
            myResult = value
        End Set
    End Property
    Public Property UVLength()
        Get
            Return myUVLength
        End Get
        Set(ByVal value)
            myUVLength = value
        End Set
    End Property
    Public Property Blank()
        Get
            Return myBlank
        End Get
        Set(ByVal value)
            myBlank = value
        End Set
    End Property
    Public ReadOnly Property Json()
        Get
            Dim DATAFObj As New JObject
            Dim DATAFArr As New JArray
            DATAFObj.Add("Samplename", myMethodId)
            Select Case myOnmicSn
                Case 61
                    DATAFObj.Add("pH", myResult)
                Case 62
                    DATAFObj.Add("S.G", myResult)
                Case 63
                    DATAFObj.Add(UVLength & "nm", myResult)
                    DATAFObj.Add("blank", myBlank)
            End Select
            DATAFArr.Add(DATAFObj)
            Dim ITEMObj As New JObject
            Dim ITEMArr As New JArray
            ITEMObj.Add("sn", myItemSn)
            ITEMObj.Add("setDate", mySetDate)
            ITEMObj.Add("dataF", DATAFArr)
            ITEMArr.Add(ITEMObj)
            Dim JSONObj As New JObject
            JSONObj.Add("mac", "C4:00:AD:07:44:7C")
            JSONObj.Add("bottleId", myBottleId)
            JSONObj.Add("preprocessDate", myPreprocessDate)
            JSONObj.Add("analysisEndDate", myAnalysisEndDate)
            JSONObj.Add("items", ITEMArr)
            Return "json=" & JSONObj.ToString
        End Get
    End Property
    Public Property Status()
        Get
            Return myStatus
        End Get
        Set(ByVal value)
            myStatus = value
        End Set
    End Property
End Class
