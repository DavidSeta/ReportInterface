Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic

Public Class clMSNFile
    Private Files As String
    Private DirectoryNamesAsRecID As Boolean = False
    'Added 8/10/2015 for ProductPartitionReporting
    Private nColumnImp As Integer = -1
    Private nColumnPG As Integer = -1
    Private nColumnType As Integer = -1
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
    Public Sub Execute(ByVal oClient As Clients, ByVal bFlexEOM As Boolean)
        Dim sDir As String = GetMyDir(oClient)
        Dim oDir As New DirectoryInfo(sDir)
        Dim oFile As FileInfo
        Dim FileDate As Date

        If Directory.Exists(sDir) = False Then Directory.CreateDirectory(sDir)
        For Each oFile In oDir.GetFiles
            FileDate = GetFileDate(oFile.Name)

            Log(oClient.LogName(True) & "File: " & oFile.Name)
            If ProcessFile(oFile, oClient, FileDate, bFlexEOM) Then MoveFile(oFile)
            Application.DoEvents()
        Next

    End Sub
#End Region 'Public Sub Execute

#Region " Public Sub Execute EOM"
    Public Function ExecuteEOM(ByVal oClient As Clients) As String
        Dim sRowDeleted As String
        Dim sMSNAcctID As String = oClient.MSN.AccountID
        Dim sCustomerID As String = oClient.CustID
        Dim sCustomerName As String = oClient.Name

        'Instantiate and set fields for use in SQL DELETE FROM PERFDATA
        Dim oDB As New clDatabase
        oDB.cs = ConnectionString
        oDB.EngineID = 3
        oDB.CustomerID = sCustomerID
        oDB.CustomerName = sCustomerName
        oDB.sEOMDate = oClient.MSN.StartDate
        oDB.eEOMDate = oClient.MSN.EndDate
        sRowDeleted = oDB.DeletePerfdata()

        Return sRowDeleted
    End Function
#End Region 'Public Sub Execute EOM

#Region " Process CSV File (Default) "
    Private Function Processfile_CSV(ByVal sFile As String, ByVal oClient As Clients, ByVal FileDate As Date, ByVal sDelim As Char, ByVal bFlexEOM As Boolean) As Boolean
        Debug.Print("Processing CSV")

        Dim stPG() As String
        Dim strPG As String
        Dim sKeywordLine As New StringBuilder()
        Dim sPG As String
        Dim sExtractPG As String
        Dim nCnt As Integer

        Dim sMSNAcctID As String = oClient.MSN.AccountID

        Dim s() As String
        Dim sCustomerID As String = oClient.CustID
        Dim sCustomerName As String = oClient.Name

        Dim sHeader(11) As String
        Dim nCount As Integer

        Dim bPGSkip As Boolean
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

        'Get report column names - convert into SQL PerfData table column names
        oDB.Columns = ParseColumns(sHeader(11), sDelim)
        oDB.DataTypes = _ColTypes
        Dim intImp As Integer
        Dim nLine As Long = 0
        Dim sLine As String
        Dim conn As New SqlConnection(ConnectionString)
        Dim cmd As New SqlCommand
        conn.Open()
        cmd.Connection = conn
        cmd.CommandTimeout = 0

        'Read each deatil line - filter out when Imp = 0 and new ProductPartitionReport.PartitionType = SubDivision
        While Not fs.EndOfStream
            sLine = fs.ReadLine
            nLine += 1
            RaiseEvent NewLine(nLine.ToString)
            sPG = ""
            sExtractPG = ""
            sKeywordLine.Length = 0
            bPGSkip = False
            'ParseColumns above will flag as ProductGroup Report Column via nColumnPG
            If sLine.Length > 0 Then
                oDB.Data = ParseMe(sLine, sDelim)
                If oDB.IsdataValid Then
                    'Determine if we have Impressions Data > 0
                    If oDB.Data(nColumnImp).Length > 0 Then
                        If IsNumeric(oDB.Data(nColumnImp)) Then
                            intImp = Integer.Parse(oDB.Data(nColumnImp))
                        Else
                            intImp = 0
                        End If
                    Else
                        intImp = 0
                    End If
                    If intImp > 0 Then

                        'New logic 8/10/2015 to test for ProductGroup and skipping all sub-divisions types in order to properly balance
                        If nColumnPG <> -1 Then
                            sPG = oDB.Data(nColumnType).Trim

                            If sPG.ToLower <> "unit" Then
                                bPGSkip = True
                            Else
                                sPG = ""
                                sKeywordLine.Length = 0
                                bPGSkip = False
                                'We have a unit record, now extract out the product group elements and resolve item for writing to DB
                                s = ParsePG(oDB.Data(nColumnPG).Trim, "\")
                                nCnt = 0
                                ReDim stPG(nCnt)
                                For nCount = s.GetLowerBound(0) To s.GetUpperBound(0)
                                    ReDim Preserve stPG(nCount)
                                    sPG = s(nCount)
                                    sPG = FindStr(sPG, "=")
                                    If sPG = "*" Or sPG.Length = 0 Then
                                    Else
                                        stPG(nCount) = sPG.Trim
                                    End If
                                Next
                                'Now update the columns data with extracted values
                                For nCnt = stPG.GetLowerBound(0) To stPG.GetUpperBound(0)
                                    If Not stPG(nCnt) Is Nothing Then
                                        If nCnt = stPG.GetUpperBound(0) Then
                                            sKeywordLine.Append(" Item = " & stPG(nCnt).Trim)
                                        Else
                                            sKeywordLine.Append(s(nCnt).Trim & "|")
                                        End If

                                    End If
                                Next
                                oDB.Data(nColumnPG) = sKeywordLine.ToString
                            End If
                        End If

                        'Have data ready for update/add to PerfData
                        If Not bPGSkip Then
                            If nColumnDate >= 0 Then
                                oDB.myDate = CDate(oDB.Data(nColumnDate))
                            End If
                            If bFlexEOM = False Then
                                oDB.UpdateSQL(oClient.IsArbitrage, cmd)
                            Else
                                oDB.UpdateSQLFlexEOM(oClient.IsArbitrage, cmd)
                            End If
                            Application.DoEvents()
                        End If

                    End If  'intImp > 0
                End If      'oDB.IsdataValid 
            End If          'sLine.Length > 0
        End While           'Read all input file line items

        'cmd.CommandText = "exec ArbitragePriceAdjust "
        'cmd.ExecuteScalar()

        conn.Close()
        RaiseEvent NewLine("")

        fs.Close()
        fs.Dispose()

        Return True
    End Function
#End Region 'Process CSV File (Default)

#Region " Process TSV File "
    Private Function Processfile_TSV(ByVal sFile As String, ByVal oClient As Clients, ByVal FileDate As Date, ByVal bFlexEOM As Boolean) As Boolean
        Debug.Print("Processing TSV")

        Return True
    End Function
#End Region 'Process TSV File

#Region " Process XML File "
    Private Function Processfile_XML(ByVal sFile As String, ByVal oClient As Clients, ByVal FileDate As Date, ByVal bFlexEOM As Boolean) As Boolean
        Debug.Print("Processing XML")

        Return True
    End Function
#End Region 'Process XML File

#Region " Private Methods "

    Private Function ParseMe(ByVal sLine As String, ByVal sDelim As Char) As String()
        Dim s() As String
        Dim sTemp As String = sLine

        sLine = sLine.Replace(sDelim & "," & sDelim, "|")     'Change the parse character
        sLine = sLine.Substring(1, sLine.Length - 2)            'Remove 1st and last qoute
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

    Private Function ProcessFile(ByVal oFile As FileInfo, ByVal oClient As Clients, ByVal FileDate As Date, ByVal bFlexEOM As Boolean) As Boolean
        Select Case oFile.Extension.ToUpper
            Case ".CSV"
                Return Processfile_Check(oFile.FullName, oClient, FileDate, bFlexEOM)

            Case ".TSV"
                Return Processfile_TSV(oFile.FullName, oClient, FileDate, bFlexEOM)

            Case ".XML"
                Return Processfile_XML(oFile.FullName, oClient, FileDate, bFlexEOM)

            Case Else
                MsgBox("Don't Know what to do with a " & oFile.Extension & " file.")

        End Select

        Return False
    End Function

    Private Function Processfile_Check(ByVal sfile As String, ByVal oClient As Clients, ByVal FileDate As Date, ByVal bFlexEOM As Boolean) As Boolean
        'Check if Tab or Comma Seperated and Run correct process
        Dim sDelim As Char
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
            sDelim = Chr(34)
            Return Processfile_CSV(sfile, oClient, FileDate, sDelim, bFlexEOM)
        Else
            sDelim = Chr(9)
            Return Processfile_CSV(sfile, oClient, FileDate, sDelim, bFlexEOM)
        End If
    End Function

    Private Function ParseColumns(ByVal sLine As String, ByVal sDelim As Char) As String()
        Dim sColumns() As String = ParseMe(sLine, sDelim)
        Dim nCount As Integer

        'Find Date and ProductGroup Column
        nColumnDate = -1
        nColumnPG = -1
        nColumnImp = -1
        nColumnType = -1
        For nCount = sColumns.GetLowerBound(0) To sColumns.GetUpperBound(0)
            'Added Aug 10, 2015 for ProductPartitionReport - only sum for *
            If sColumns(nCount).ToLower = "impressions" Then
                nColumnImp = nCount
            End If
            If sColumns(nCount).ToLower = "productgroup" Then
                nColumnPG = nCount
            End If
            If sColumns(nCount).ToLower = "productpartitiontype" Then
                nColumnType = nCount
            End If
            If sColumns(nCount).ToLower = "gregoriandate" Then
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

    Public Function UpdateRemarketing_Keywords(ByVal psDate As Date, ByVal peDate As Date) As String
        Debug.Print("Update MSN Remarketing Keyword")
        Dim sSDate As String
        Dim sEDate As String
        sSDate = psDate
        sEDate = peDate

        Dim conn As New SqlConnection(ConnectionString)
        Dim cmd As New SqlCommand
        conn.Open()
        cmd.CommandTimeout = 0
        cmd.Connection = conn
        'Stored Procedure to populate PerfDataArbitrage.ModifiedCost based on Customers.ArbitrageMarkUP
        cmd.CommandText = "BoomUserList"
        cmd.CommandType = CommandType.StoredProcedure

        cmd.Parameters.AddWithValue("@charSDate", sSDate)
        cmd.Parameters.AddWithValue("@charEDate", sEDate)
        cmd.CommandTimeout = 0
        Dim intCount As Integer
        intCount = cmd.ExecuteNonQuery()

        conn.Close()
        Return intCount.ToString
    End Function
    Private Function ParsePG(ByVal sLine As String, ByVal sDelim As String) As String()
        Dim s() As String
        Dim sTemp As String = sLine

        sLine = sLine.Replace(sDelim, "|")     'Change the parse character
        sLine = sLine.Substring(1, sLine.Length - 2)            'Remove 1st and last qoute
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
    Private Function FindStr(ByVal sLine As String, ByVal sDelim As Char) As String
        Dim s As String
        s = sLine
        If s.IndexOf(sDelim) <> -1 Then
            s = s.Substring(s.LastIndexOf("=") + 1)
        Else
            If s.IndexOf("$t38") <> -1 Then
                s = s.Substring(s.IndexOf("$t38") + 1)
                s = s.Substring(3)
            End If
        End If

        Return s.Trim

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
            'FileDate = GetFileDate(oFile.Name)

            'If FileDate >= StartDate Then
            '    If FileDate <= EndDate Then
            '        Dim sFile As String = sDir & oFile.Name
            '        If File.Exists(sFile) Then File.Delete(sFile)
            '        oFile.MoveTo(sFile)
            '        Application.DoEvents()
            '    End If
            'End If

            'January Reset
            If (oFile.Name.IndexOf("2010") <> -1) Then
                'If oFile.Name.ToString = MyFile Then

                Dim sFile As String = sDir & oFile.Name
                If File.Exists(sFile) Then File.Delete(sFile)
                oFile.MoveTo(sFile)
                Application.DoEvents()

            End If
            Application.DoEvents()
        Next
    End Sub

    Public Sub ResetFilesEOM(ByVal Client As Clients, ByVal StartDate As Date, ByVal EndDate As Date)
        Dim sDir As String = GetMyDir(Client)
        Dim sDirP As String = sDir & "Processed\"

        Dim oDir As New DirectoryInfo(sDirP)
        Dim oFile As FileInfo

        Dim FileDate As Date

        If Directory.Exists(sDir) = False Then Directory.CreateDirectory(sDir)
        If Directory.Exists(sDirP) = False Then Directory.CreateDirectory(sDirP)

        For Each oFile In oDir.GetFiles
            'FileDate = GetFileDate(oFile.Name)

            'If FileDate >= StartDate Then
            '    If FileDate <= EndDate Then
            '        Dim sFile As String = sDir & oFile.Name
            '        If File.Exists(sFile) Then File.Delete(sFile)
            '        oFile.MoveTo(sFile)
            '        Application.DoEvents()
            '    End If
            'End If

            'January Reset
            If (oFile.Name.IndexOf("2010") <> -1) Then
                'If oFile.Name.ToString = MyFile Then

                Dim sFile As String = sDir & oFile.Name
                If File.Exists(sFile) Then File.Delete(sFile)
                oFile.MoveTo(sFile)
                Application.DoEvents()

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
