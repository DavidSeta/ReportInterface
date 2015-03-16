Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic

Public Class clMSNFile
    Private Files As String
    Private DirectoryNamesAsRecID As Boolean = False

    Private nColumnDate As Integer = -1
    Private _Columns As Hashtable = New Hashtable()
    Private _ColTypes As Hashtable = New Hashtable()
    Private ConnectionString As String

    Public Event NewLine(ByVal sLine As String)


#Region " Sub New "
    Sub New(ByVal oSettings As clSettings)
        Files = oSettings.SaveBaseLocation & "MSN\"
        ConnectionString = oSettings.ConnectionString
        DirectoryNamesAsRecID = oSettings.DirectoryNamesAsRecID

        Dim nCount As Integer

        _Columns.Clear()
        For nCount = oSettings.MSNColumns.GetLowerBound(0) To oSettings.MSNColumns.GetUpperBound(0)
            If Not IsNothing(oSettings.MSNColumns(nCount)) Then
                _Columns.Add(oSettings.MSNColumns(nCount).Name, oSettings.MSNColumns(nCount).Value)
            End If
        Next

        _ColTypes.Clear()
        For nCount = oSettings.ColumnTypes.GetLowerBound(0) To oSettings.ColumnTypes.GetUpperBound(0)
            If Not IsNothing(oSettings.ColumnTypes(nCount)) Then
                _ColTypes.Add(oSettings.ColumnTypes(nCount).Name.ToLower, oSettings.ColumnTypes(nCount).Value)
            End If
        Next

    End Sub
#End Region 'Sub New

#Region " Public Sub Execute "
    Public Sub Execute(ByVal oClient As Clients)
        Dim sDir As String = GetMyDir(oClient)
        Dim oDir As New DirectoryInfo(sDir)
        Dim oFile As FileInfo
        Dim FileDate As Date

        If Directory.Exists(sDir) = False Then Directory.CreateDirectory(sDir)
        For Each oFile In oDir.GetFiles
            FileDate = GetFileDate(oFile.Name)

            Log(oClient.LogName(True) & "File: " & oFile.Name)
            If ProcessFile(oFile, oClient, FileDate) Then MoveFile(oFile)
            Application.DoEvents()
        Next

    End Sub
#End Region 'Public Sub Execute

#Region " Process CSV File (Default) "
    Private Function Processfile_CSV(ByVal sFile As String, ByVal oClient As Clients, ByVal FileDate As Date) As Boolean
        Debug.Print("Processing CSV")

        Dim sMSNAcctID As String = oClient.MSN.AccountID


        Dim sCustomerID As String = oClient.CustID
        Dim sCustomerName As String = oClient.Name

        Dim sHeader(11) As String
        Dim nCount As Integer
        Dim fs As StreamReader = File.OpenText(sFile)
        For nCount = 1 To 11
            sHeader(nCount) = fs.ReadLine
        Next

        Dim oDB As New clDatabase
        oDB.cs = ConnectionString
        oDB.EngineID = 3
        oDB.myDate = FileDate
        oDB.CustomerID = sCustomerID
        oDB.CustomerName = sCustomerName

        oDB.Columns = ParseColumns(sHeader(11))
        oDB.DataTypes = _ColTypes

        Dim nLine As Long = 0
        Dim sLine As String
        While Not fs.EndOfStream
            sLine = fs.ReadLine
            nLine += 1
            RaiseEvent NewLine(nLine.ToString)

            If sLine.Length > 0 Then
                oDB.Data = Unquote(sLine.Split(","))
                If oDB.IsdataValid Then
                    If nColumnDate >= 0 Then
                        oDB.myDate = CDate(oDB.Data(nColumnDate))
                    End If
                    oDB.UpdateSQL()
                    Application.DoEvents()
                End If

            End If
        End While
        RaiseEvent NewLine("")

        fs.Close()
        fs.Dispose()

        Return True
    End Function
#End Region 'Process CSV File (Default)

#Region " Process TSV File "
    Private Function Processfile_TSV(ByVal sFile As String, ByVal oClient As Clients, ByVal FileDate As Date) As Boolean
        Debug.Print("Processing TSV")

        Return True
    End Function
#End Region 'Process TSV File

#Region " Process XML File "
    Private Function Processfile_XML(ByVal sFile As String, ByVal oClient As Clients, ByVal FileDate As Date) As Boolean
        Debug.Print("Processing XML")

        Return True
    End Function
#End Region 'Process XML File

#Region " Private Methods "
    Private Function Unquote(ByVal s As String()) As String()
        Dim nCount As Integer
        For nCount = s.GetLowerBound(0) To s.GetUpperBound(0)
            If s(nCount).Length > 2 Then
                s(nCount) = s(nCount).Substring(1, s(nCount).Length - 2)
            End If
        Next
        Return s
    End Function

    Private Function ProcessFile(ByVal oFile As FileInfo, ByVal oClient As Clients, ByVal FileDate As Date) As Boolean
        Select Case oFile.Extension.ToUpper
            Case ".CSV"
                Return Processfile_Check(oFile.FullName, oClient, FileDate)

            Case ".TSV"
                Return Processfile_TSV(oFile.FullName, oClient, FileDate)

            Case ".XML"
                Return Processfile_XML(oFile.FullName, oClient, FileDate)

            Case Else
                MsgBox("Don't Know what to do with a " & oFile.Extension & " file.")

        End Select

        Return False
    End Function

    Private Function Processfile_Check(ByVal sfile As String, ByVal oClient As Clients, ByVal FileDate As Date) As Boolean
        'Check if Tab or Comma Seperated and Run correct process

        Dim sLine As String = ""
        Dim fs As StreamReader = File.OpenText(sfile)

        Dim nCount As Integer
        For nCount = 1 To 11
            sLine = fs.ReadLine
        Next
        fs.Close()
        fs.Dispose()

        Dim sCSV() As String = sLine.Split(",")
        Dim sTSV() As String = sLine.Split(vbTab)

        If sCSV.GetUpperBound(0) > sTSV.GetUpperBound(0) Then
            Return Processfile_CSV(sfile, oClient, FileDate)
        Else
            Return Processfile_TSV(sfile, oClient, FileDate)
        End If
    End Function

    Private Function ParseColumns(ByVal sLine As String) As String()
        Dim sColumns() As String = Unquote(sLine.Split(","))

        Dim nCount As Integer

        nColumnDate = -1
        For nCount = sColumns.GetLowerBound(0) To sColumns.GetUpperBound(0)
            If sColumns(nCount).ToLower = "date" Then
                nColumnDate = nCount
                sColumns(nCount) = ""
            Else
                sColumns(nCount) = _Columns(sColumns(nCount))
            End If
        Next

        Return sColumns
    End Function

    Private Function GetDataTypes(ByVal Columns() As String) As String()
        Dim nCount As Integer = Columns.GetUpperBound(0)
        Dim DataTypes(nCount) As String

        For nCount = Columns.GetLowerBound(0) To Columns.GetUpperBound(0)
            If Columns(nCount).Length = 0 Then
                DataTypes(nCount) = ""
            Else
                DataTypes(nCount) = _ColTypes(Columns(nCount))
            End If
        Next

        Return DataTypes
    End Function

    Private Sub MoveFile(ByVal oFile As FileInfo)
        Dim sDir As String = oFile.DirectoryName & "\Processed\"
        Dim sFile As String = sDir & oFile.Name

        If Directory.Exists(sDir) = False Then Directory.CreateDirectory(sDir)
        If File.Exists(sFile) Then File.Delete(sFile)
        oFile.MoveTo(sFile)
    End Sub

    Private Sub RenameFile(ByVal oFile As FileInfo, ByVal sNewExt As String)
        If sNewExt = "" Then Exit Sub
        If oFile.Extension = sNewExt Then Exit Sub
        Dim sNew As String = Left(oFile.FullName, Len(oFile.FullName) - 3) & sNewExt
        Rename(oFile.FullName, sNew)
    End Sub

    Private Function GetMyDir(ByVal Client As Clients) As String
        '----------------------------------------------------------------------------------------------------------------
        ' Changed 3/14/07 so directories are now RecID's instead of names
        '----------------------------------------------------------------------------------------------------------------
        Dim sDir As String
        If DirectoryNamesAsRecID Then
            sDir = Files & Client.CustID & "\"
        Else
            sDir = Files & Client.Name & "\"
        End If
        Return sDir
    End Function

#End Region 'Private Methods

    Private Function GetFileDate(ByVal sFileName As String) As Date
        Dim sTemp As String = sFileName.Split(".")(0)
        sTemp = Right(sTemp, 10)

        Dim myFileDate As Date = CDate(Mid(sTemp, 6, 2) & "/" & Right(sTemp, 2) & "/" & Left(sTemp, 4))

        Return myFileDate
    End Function

    '----------------------------------------------------------------------------------------------------------------
    ' Added 3/14/07 so directories are now RecID's instead of names
    '----------------------------------------------------------------------------------------------------------------
    Public Sub RenameDirs(ByVal Client As Clients)
        Dim sDirOld As String
        Dim sDirNew As String
        If DirectoryNamesAsRecID Then
            sDirOld = Files & Client.Name
            sDirNew = Files & Client.CustID
        Else
            sDirOld = Files & Client.CustID
            sDirNew = Files & Client.Name
        End If
        If Directory.Exists(sDirOld) Then Directory.Move(sDirOld, sDirNew)
    End Sub

#Region " Reset DB Files Sub "
    Public Sub ResetFiles(ByVal Client As Clients, ByVal StartDate As Date, ByVal EndDate As Date)
        Dim sDir As String = GetMyDir(Client)
        Dim sDirP As String = sDir & "Processed\"

        Dim oDir As New DirectoryInfo(sDirP)
        Dim oFile As FileInfo

        Dim FileDate As Date

        If Directory.Exists(sDir) = False Then Directory.CreateDirectory(sDir)
        If Directory.Exists(sDirP) = False Then Directory.CreateDirectory(sDirP)

        For Each oFile In oDir.GetFiles
            FileDate = GetFileDate(oFile.Name)

            If FileDate >= StartDate Then
                If FileDate <= EndDate Then
                    Dim sFile As String = sDir & oFile.Name
                    If File.Exists(sFile) Then File.Delete(sFile)
                    oFile.MoveTo(sFile)
                    Application.DoEvents()
                End If
            End If
            Application.DoEvents()
        Next
    End Sub
#End Region 'Reset DB Files Sub

#Region " CheckProcessed "
    Public Function CheckNotProcessed(ByVal Clients() As Clients, ByVal CheckDate As Date, ByVal bCheckProcessedDir As Boolean) As String()
        Dim sBad(0) As String
        sBad(0) = "(MSN) Date: " & CheckDate.ToString("yyyy_MM_dd")
        Dim Index As Integer = 0

        For Each client As Clients In Clients
            If client.MSN.Has_AccountID Then
                Dim sDir As String = GetMyDir(client)
                If bCheckProcessedDir = True Then sDir = sDir & "Processed\"

                Dim sFileName1 As String = sDir & "M_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".csv"
                Dim sFileName2 As String = sDir & "MD_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".csv"
                Dim sFileName3 As String = sDir & "MM_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".csv"

                Dim bExists As Boolean = False
                If File.Exists(sFileName1) Then bExists = True
                If File.Exists(sFileName2) Then bExists = True
                If File.Exists(sFileName3) Then bExists = True

                If bExists = False Then
                    Index = sBad.GetUpperBound(0) + 1
                    ReDim Preserve sBad(Index)
                    sBad(Index) = "MSN MISSING FOR (" & client.CustID.ToString & ") " & client.Name
                End If
            End If
        Next

        If Index = 0 Then sBad(0) = ""
        Return sBad
    End Function
#End Region 'CheckProcessed

End Class
