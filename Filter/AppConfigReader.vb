Public Class AppConfigReader
    Private Shared aSettingsReader As New System.Configuration.AppSettingsReader

    Private Shared strfPath As String = aSettingsReader.GetValue("filePath", GetType(String))
    Private Shared strName As String = aSettingsReader.GetValue("Name", GetType(String))
    Private Shared strdPath As String = aSettingsReader.GetValue("defaultPath", GetType(String))

    Public Shared ReadOnly Property filePath() As String
        Get
            Return strfPath
        End Get
    End Property

    Public Shared ReadOnly Property Name() As String
        Get
            Return strName
        End Get
    End Property

    Public Shared ReadOnly Property defaultPath() As String
        Get
            Return strdPath
        End Get
    End Property
End Class
