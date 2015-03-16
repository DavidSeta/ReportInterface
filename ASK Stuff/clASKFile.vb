Imports System.IO
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic

Public Class clASKFile
    Private Files As String
    Private DirectoryNamesAsRecID As Boolean = False

    Private nColumnDate As Integer = -1
    Private sColumnNames() As String
    Private sColumns() As String
    Private _Columns As Hashtable = New Hashtable()
    Private _ColTypes As Hashtable = New Hashtable()
    Private ConnectionString As String

    Public Event NewLine(ByVal sLine As String)

#Region " Sub New "
    '================================================================================
    'New(): Instantiation of object - Invoked from frmMain()
    'Will create two HashTable objects from settings.xml/oSettings.ASKColumns
    'One for Column names, _Columns, and another for Column Types, _ColTypes
    '================================================================================
    Sub New(ByVal oSettings As clSettings)
        Files = oSettings.SaveBaseLocation & "ASK\"
        ConnectionString = oSettings.ConnectionString
        DirectoryNamesAsRecID = oSettings.DirectoryNamesAsRecID

        Dim nCount As Integer
        _Columns.Clear()
        For nCount = oSettings.ASKColumns.GetLowerBound(0) To oSettings.ASKColumns.GetUpperBound(0)
            If Not IsNothing(oSettings.ASKColumns(nCount)) Then
                _Columns.Add(oSettings.ASKColumns(nCount).Name, oSettings.ASKColumns(nCount).Value)
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
    '================================================================================
    'Execute(oClient): Invoked from frmMain() for nStep=50
    'File work for response object from report request API
    '================================================================================
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

#Region " Process CSV File "
    '================================================================================
    'Processfile_CSV(): Process file from API and load into DB
    '================================================================================
    Private Function Processfile_CSV(ByVal sFile As String, ByVal oClient As Clients, ByVal FileDate As Date) As Boolean
        Debug.Print("Processing CSV")
        Dim sASKAcctID As String = oClient.ASK.AccountID

        Dim sCustomerID As String = oClient.CustID
        Dim sCustomerName As String = oClient.Name

        Dim sHeader As String

        Dim fs As StreamReader = File.OpenText(sFile)

        'Get Header row from file
        sHeader = fs.ReadLine

        'Bld DB values for add/update
        Dim oDB As New clDatabase
        oDB.cs = ConnectionString
        oDB.EngineID = 5
        oDB.myDate = FileDate
        oDB.CustomerID = sCustomerID
        oDB.CustomerName = sCustomerName

        'Mappings for Column names used in API build file download to the DB column names
        oDB.Columns = ParseColumns(sHeader)
        oDB.DataTypes = _ColTypes

        Dim nLine As Long = 0
        Dim sLine As String
        Dim conn As New SqlConnection(ConnectionString)
        Dim cmd As New SqlCommand
        conn.Open()
        cmd.Connection = conn
        While Not fs.EndOfStream
            sLine = fs.ReadLine
            nLine += 1
            RaiseEvent NewLine(nLine.ToString)

            If sLine.Length > 0 Then
                oDB.Data = sLine.Split(",")
                If oDB.IsdataValid Then
                    If nColumnDate >= 0 Then
                        oDB.myDate = CDate(oDB.Data(nColumnDate))
                    End If
                    oDB.UpdateSQL(oClient.IsArbitrage, cmd)
                    Application.DoEvents()
                End If

            End If
        End While
        'cmd.CommandText = "exec ArbitragePriceAdjust "
        'cmd.ExecuteScalar()

        conn.Close()
        conn.Dispose()
        cmd.Dispose()
        RaiseEvent NewLine("")

        fs.Close()
        fs.Dispose()

        Return True
    End Function
#End Region 'Process CSV File

#Region " Process TSV File "
    Private Function Processfile_TSV(ByVal sFile As String, ByVal oClient As Clients, ByVal FileDate As Date) As Boolean
        Debug.Print("Processing TSV")

        Return True
    End Function
#End Region 'Process TSV File

#Region " Process XML File (Default)"
    Private Function Processfile_XML(ByVal sFile As String, ByVal oClient As Clients, ByVal FileDate As Date) As Boolean
        Debug.Print("Processing XML")

        Dim sASKAcctID As String = oClient.ASK.AccountID
        Dim bHaveData As Boolean
        Dim sCustomerID As String = oClient.CustID
        Dim sCustomerName As String = oClient.Name

        Dim oDB As New clDatabase
        oDB.cs = ConnectionString
        oDB.EngineID = 2
        oDB.myDate = FileDate
        oDB.CustomerID = sCustomerID
        oDB.CustomerName = sCustomerName

        oDB.Columns = sColumns
        oDB.DataTypes = _ColTypes

        Dim nCount As Integer

        Dim reader As New XmlTextReader(sFile)
        reader.WhitespaceHandling = WhitespaceHandling.None

        Dim nLine As Long = 0
        Dim sData() As String
        'Loop through the XML document and process once we find a "row"
        Dim conn As New SqlConnection(ConnectionString)
        Dim cmd As New SqlCommand
        conn.Open()
        cmd.Connection = conn
        While reader.Read()
            If reader.Name.ToLower = "row" Then

                'Increament record counter and show on windows form
                nLine += 1
                RaiseEvent NewLine(nLine.ToString)

                ReDim sData(sColumnNames.GetUpperBound(0))
                bHaveData = False
                For nCount = 0 To sColumnNames.GetUpperBound(0)
                    'Addded for clients w/o conversions
                    If ((Not IsNothing(reader.GetAttribute(sColumnNames(nCount)))) Or (sColumnNames(nCount).ToLower = "kwsite")) Then
                        bHaveData = True
                        'Move value to array
                        sData(nCount) = reader.GetAttribute(sColumnNames(nCount))
                        'Update if no value - logic for the clients w/o converstion data

                        If sColumnNames(nCount).ToLower = "cost" Then
                            Dim nTemp As Double = reader.GetAttribute(sColumnNames(nCount))
                            nTemp = nTemp / 1000000
                            sData(nCount) = nTemp.ToString
                        End If
                        If sColumnNames(nCount).ToLower = "cpc" Then
                            Dim nTemp As Double = reader.GetAttribute(sColumnNames(nCount))
                            nTemp = nTemp / 1000000
                            sData(nCount) = nTemp.ToString
                        End If

                        '------------------------------------------------------------------------------------------------------
                        ' 3-13-07 Fix for different keyword names in Google downloads (keyword) (kwSite)
                        '------------------------------------------------------------------------------------------------------
                        If sColumnNames(nCount).ToLower = "kwsite" Then
                            If IsNothing(reader.GetAttribute(sColumnNames(nCount))) Then
                                sData(nCount) = reader.GetAttribute("keyword")
                            End If
                        End If

                        If sColumnNames(nCount).ToLower = "keyword" Then
                            If IsNothing(reader.GetAttribute(sColumnNames(nCount))) Then
                                sData(nCount) = reader.GetAttribute("kwSite")
                            End If
                        End If
                        '------------------------------------------------------------------------------------------------------
                    Else
                        sData(nCount) = "0"
                    End If

                Next    'For nCount

                'Must have data to proceed.
                If ((bHaveData) And (sData(1) <> "0")) Then
                    oDB.Data = sData
                    If Not IsNothing(oDB.Data(nColumnDate)) Then
                        If nColumnDate >= 0 Then
                            oDB.myDate = CDate(oDB.Data(nColumnDate))
                        End If

                        oDB.UpdateSQL(oClient.IsArbitrage, cmd)
                        Application.DoEvents()
                    End If
                End If 'If we had any data to insert/update into our DB

            End If   'If reader.Name.ToLower = "row"
        End While    'While reader.Read()
        'cmd.CommandText = "exec ArbitragePriceAdjust "
        'cmd.ExecuteScalar()

        conn.Close()
        conn.Dispose()
        cmd.Dispose()

        RaiseEvent NewLine("")
        reader.Close()
        Return True
    End Function
#End Region 'Process XML File (Default)

#Region " Private Methods "
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

        sLine = fs.ReadLine

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
    Private Function ParseColumns(ByVal sLine As String) As String()
        Dim sColumns() As String = sLine.Split(",")
        Dim nCount As Integer
        For nCount = sColumns.GetLowerBound(0) To sColumns.GetUpperBound(0)
            sColumns(nCount) = _Columns(sColumns(nCount))
        Next

        Return sColumns
    End Function
    Private Function ParseMe(ByVal sLine As String) As String()
        Dim s() As String
        Dim sTemp As String = sLine

        sLine = sLine.Replace(",", "|")                         'Change the parse character from a "," to a "|"
        sLine = sLine.Substring(0, sLine.Length - 1)            'Remove 1st and last qoute
        sLine = sLine.Replace(",", "")                          'Remove any comma seperators remaining
        s = sLine.Split("|")                                    'Now parse it

        Dim nCount As Integer
        For nCount = 0 To s.GetUpperBound(0)
            If s(nCount).Contains(",") Then
                If IsNumeric(s(nCount)) Then
                    s(nCount).Replace(",", "")                      'Remove comma if this is a numeric string
                End If
            End If
        Next

        Return s
    End Function
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
#End Region 'Private Methods

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
        sBad(0) = "(ASK) Date: " & CheckDate.ToString("yyyy_MM_dd")
        Dim Index As Integer = 0

        For Each client As Clients In Clients
            If client.ASK.Has_AccountID Then

                Dim sDir As String = GetMyDir(client)
                If bCheckProcessedDir = True Then sDir = sDir & "Processed\"

                Dim sFileName1 As String = sDir & "A_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".csv"
                Dim sFileName2 As String = sDir & "AD_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".csv"
                Dim sFileName3 As String = sDir & "AM_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".csv"

                Dim bExists As Boolean = False
                If File.Exists(sFileName1) Then bExists = True
                If File.Exists(sFileName2) Then bExists = True
                If File.Exists(sFileName3) Then bExists = True

                If bExists = False Then
                    Index = sBad.GetUpperBound(0) + 1
                    ReDim Preserve sBad(Index)
                    sBad(Index) = "ASK MISSING FOR (" & client.CustID.ToString & ") " & client.Name
                End If
            End If
        Next

        If Index = 0 Then sBad(0) = ""
        Return sBad
    End Function
#End Region 'CheckProcessed
End Class
