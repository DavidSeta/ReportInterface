Imports System.Data.SqlClient
Imports System.Data.Sql

Public Class clEngineNames
    Dim osettings As clSettings
    Dim sCnstr As String

    Dim _HashEngines As New Hashtable()

    Dim _EngineName As String
    Dim _EngineID As Integer
 

    '===================================================================================================================
    'New() - Instantiate Object 
    'Engine authentication work 
    '===================================================================================================================
    Sub New()

        osettings = New clSettings()
        sCnstr = osettings.ConnectionString
        Dim ssql As String = " SELECT EngineID, EngineName"
        ssql &= " FROM SearchEngines ORDER BY Seq"

        Dim cnn As SqlConnection = Nothing
        cnn = New SqlConnection()

        Dim dr As SqlDataReader = Nothing


        Dim cmd As SqlCommand = Nothing
        cmd = New SqlCommand()

        cnn.ConnectionString = sCnstr
        cnn.Open()

        cmd = New SqlCommand
        cmd.Connection = cnn
        cmd.CommandText = ssql
        cmd.CommandType = CommandType.Text

        'Instantiate DataReader
        dr = cmd.ExecuteReader()
        Dim sEngineName As String
        Dim intEngineId As Integer

        While dr.Read()
            sEngineName = dr(1).ToString.ToLower()
            intEngineId = dr(0)
            _HashEngines.Add(sEngineName, intEngineId)
        End While
        If Not (cmd Is Nothing) Then
            cmd.Dispose()
        End If
        If Not (dr Is Nothing) Then
            dr.Close()
        End If
        If Not (cnn Is Nothing) Then
            cnn.Close()
        End If
    End Sub
    Public ReadOnly Property HashEngines() As Hashtable
        Get
            Return _HashEngines
        End Get
    End Property

End Class
