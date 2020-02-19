Public Class Platform
    Dim myStatus As String = Nothing
    Public Property Status()
        Get
            Return myStatus
        End Get
        Set(ByVal value)
            myStatus = value
        End Set
    End Property
End Class
