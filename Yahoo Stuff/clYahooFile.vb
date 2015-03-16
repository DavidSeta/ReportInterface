Imports System.IO
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic

Public Class clYahooFile
    Private Const IDS_ContentMatchID As String = "24"
    Private Const IDS_SponsoredSearchID As String = "21"

    Private Files As String
    Private DirectoryNamesAsRecID As Boolean = False

    Private nFieldAcctID As Integer
    Private nFieldAcctName As Integer
    Private nFieldTacticID As Integer

    Private _ColumnsKeyword As Hashtable = New Hashtable()
    Private _ColumnsAdGroup As Hashtable = New Hashtable()
    Private _ColTypes As Hashtable = New Hashtable()
    Private ConnectionString As String

    Public Event NewLine(ByVal sLine As String)

#Region " Sub New "
    Sub New(ByVal oSettings As clSettings)
        Files = oSettings.SaveBaseLocation & "Yahoo\"
        ConnectionString = oSettings.ConnectionString
        DirectoryNamesAsRecID = oSettings.DirectoryNamesAsRecID

        Dim nCount As Integer

        _ColumnsKeyword.Clear()
        For nCount = oSettings.YahooColumnsKeyword.GetLowerBound(0) To oSettings.YahooColumnsKeyword.GetUpperBound(0)
            If Not IsNothing(oSettings.YahooColumnsKeyword(nCount)) Then
                _ColumnsKeyword.Add(oSettings.YahooColumnsKeyword(nCount).Name, oSettings.YahooColumnsKeyword(nCount).Value)
            End If
        Next

        _ColumnsAdGroup.Clear()
        For nCount = oSettings.YahooColumnsAdGroup.GetLowerBound(0) To oSettings.YahooColumnsAdGroup.GetUpperBound(0)
            If Not IsNothing(oSettings.YahooColumnsAdGroup(nCount)) Then
                _ColumnsAdGroup.Add(oSettings.YahooColumnsAdGroup(nCount).Name, oSettings.YahooColumnsAdGroup(nCount).Value)
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
    '===============================================================================================================
    'Execute(oClient) Called via frmMain.UpdateDB_Yahoo()
    'Perform directory work and import of file data into the data base.
    'Yahoo logic added for catching incomplete file data.  March 19, 2008
    '===============================================================================================================
    Public Sub Execute(ByVal oClient As Clients)
        Dim sDir As String = GetMyDir(oClient)
        Dim oDir As New DirectoryInfo(sDir)
        Dim oFile As FileInfo
        Dim FileDate As Date

        If Directory.Exists(sDir) = False Then Directory.CreateDirectory(sDir)
        For Each oFile In oDir.GetFiles
            FileDate = GetFileDate(oFile.Name)
            Log(oClient.LogName(True) & "File: " & oFile.Name)
            If ProcessFile(oFile, oClient, FileDate) Then
                MoveFile(oFile)
            Else
                Log(oClient.LogName(True) & "Incomplete File: " & oFile.Name)
                oFile.Delete()
            End If

            Application.DoEvents()
        Next

    End Sub
#End Region 'Public Sub Execute

#Region " Process CSV File "
    Private Function Processfile_CSV(ByVal sFile As String, ByVal oclient As Clients, ByVal FileDate As Date, ByVal nCountLimit As Integer) As Boolean
        Debug.Print("Processing CSV")
        Dim nCount As Integer
        Dim sHeader As String = ""
        Dim bResult As Boolean = False

        Dim fs As StreamReader = File.OpenText(sFile)
        For nCount = 1 To nCountLimit
            sHeader = fs.ReadLine
        Next

        fs.Close()
        fs.Dispose()

        Dim sColumns() As String = sHeader.Split(",")
        Dim bTactic As Boolean = False

        If sColumns.GetUpperBound(0) >= 2 Then
            For nCount = 0 To sColumns.GetUpperBound(0)
                If sColumns(2).ToUpper = "TACTIC ID" Then bTactic = True
            Next
        End If

        If bTactic Then
            Debug.Print("Ad Group File")
            bResult = ProcessAdGroupfile_CSV(sFile, oclient, FileDate, nCountLimit)
        Else
            Debug.Print("Keyword File")
            bResult = ProcessKeywordfile_CSV(sFile, oclient, FileDate, nCountLimit)
        End If

        Return bResult
        Return True
    End Function
    Private Function ProcessKeywordfile_CSV(ByVal sFile As String, ByVal oclient As Clients, ByVal FileDate As Date, ByVal nCountLimit As Integer) As Boolean
        Debug.Print("Processing Keyword CSV")
        Dim nCntLimit As Integer = nCountLimit + 1

        Dim sYahooAcctID As String = oclient.Yahoo.AccountID
        Dim sYahooSubAcctID As String = oclient.Yahoo.SubAccountID

        Dim sCustomerID As String = oclient.CustID
        Dim sCustomerName As String = oclient.Name

        Dim sHeader(nCntLimit) As String
        Dim nCount As Integer
        Dim fs As StreamReader = File.OpenText(sFile)
        For nCount = 1 To nCntLimit
            sHeader(nCount) = fs.ReadLine
        Next

        Dim oDB As New clDatabase
        oDB.cs = ConnectionString
        oDB.EngineID = 1
        oDB.myDate = FileDate

        oDB.CustomerID = sCustomerID
        oDB.CustomerName = sCustomerName

        oDB.Columns = ParseKeywordColumnsCSV(sHeader(nCountLimit))
        oDB.DataTypes = _ColTypes

        oDB.FieldWithYahooAccountID = nFieldAcctID
        oDB.FieldWithYahooAccountName = nFieldAcctName
        oDB.YahooAccount = sYahooAcctID

        Dim nLine As Long = 0
        Dim sLine As String
        Dim conn As New SqlConnection(ConnectionString)
        Dim cmd As New SqlCommand
        conn.Open()
        cmd.Connection = conn
        If sYahooAcctID = "214045" Then
            'Yahoo Multi Account
            oDB.MemberOfMaster = True
            While Not fs.EndOfStream
                sLine = fs.ReadLine
                nLine += 1
                RaiseEvent NewLine(nLine.ToString)

                oDB.Data = sLine.Split(",")
                If oDB.Data.GetUpperBound(0) = oDB.Columns.GetUpperBound(0) Then
                    oDB.YahooAccount = oDB.Data(nFieldAcctID) 'Set for each line
                    oDB.UpdateSQL(oclient.IsArbitrage, cmd)
                End If
                Application.DoEvents()
            End While
        Else
            'Yahoo Single Account
            While Not fs.EndOfStream
                sLine = fs.ReadLine
                nLine += 1
                RaiseEvent NewLine(nLine)

                oDB.Data = sLine.Split(",")
                oDB.UpdateSQL(oclient.IsArbitrage, cmd)
                Application.DoEvents()
            End While
        End If
        cmd.CommandText = "exec ArbitragePriceAdjust "
        cmd.ExecuteScalar()

        conn.Close()
        RaiseEvent NewLine("")

        fs.Close()
        fs.Dispose()

        Return True
    End Function

    Private Function ProcessAdGroupfile_CSV(ByVal sFile As String, ByVal oclient As Clients, ByVal FileDate As Date, ByVal nCountLimit As Integer) As Boolean
        Debug.Print("Processing AdGroup CSV")
        Dim nCntLimit As Integer = nCountLimit + 1

        Dim sYahooAcctID As String = oclient.Yahoo.AccountID
        Dim sYahooSubAcctID As String = oclient.Yahoo.SubAccountID

        Dim sCustomerID As String = oclient.CustID
        Dim sCustomerName As String = oclient.Name

        Dim sHeader(nCntLimit) As String
        Dim nCount As Integer
        Dim fs As StreamReader = File.OpenText(sFile)
        For nCount = 1 To nCntLimit
            sHeader(nCount) = fs.ReadLine
        Next

        Dim sData() As String = Nothing
        Dim Index As Integer = 0

        Dim oDB As New clDatabase
        oDB.cs = ConnectionString
        oDB.EngineID = 1
        oDB.myDate = FileDate

        oDB.CustomerID = sCustomerID
        oDB.CustomerName = sCustomerName

        oDB.Columns = ParseAdGroupColumnsCSV(sHeader(nCountLimit))
        oDB.DataTypes = _ColTypes

        oDB.FieldWithYahooAccountID = nFieldAcctID
        oDB.FieldWithYahooAccountName = nFieldAcctName
        oDB.YahooAccount = sYahooAcctID

        Dim nLine As Long = 0
        Dim sLine As String
        Dim conn As New SqlConnection(ConnectionString)
        Dim cmd As New SqlCommand
        conn.Open()
        cmd.Connection = conn
        If sYahooAcctID = "214045" Then
            'Yahoo Multi Account
            oDB.MemberOfMaster = True
            While Not fs.EndOfStream
                sLine = fs.ReadLine
                nLine += 1
                RaiseEvent NewLine(nLine.ToString)

                sData = sLine.Split(",")
                Index = sData.GetUpperBound(0) + 1
                ReDim Preserve sData(Index)
                sData(Index) = "Content Match"

                If sData.GetUpperBound(0) = oDB.Columns.GetUpperBound(0) Then
                    If nFieldTacticID >= 0 Then
                        If sData(nFieldTacticID) = IDS_ContentMatchID Then
                            'Updata DB
                            oDB.YahooAccount = sData(nFieldAcctID) 'Set for each line
                            oDB.Data = sData
                            oDB.UpdateSQL(oclient.IsArbitrage, cmd)
                        End If
                    End If
                End If
                Application.DoEvents()
            End While
        Else
            'Yahoo Single Account
            While Not fs.EndOfStream
                sLine = fs.ReadLine
                nLine += 1
                RaiseEvent NewLine(nLine)

                sData = sLine.Split(",")
                Index = sData.GetUpperBound(0) + 1
                ReDim Preserve sData(Index)
                sData(Index) = "Content Match"

                If sData.GetUpperBound(0) = oDB.Columns.GetUpperBound(0) Then
                    If nFieldTacticID >= 0 Then
                        If sData(nFieldTacticID) = IDS_ContentMatchID Then
                            'Updata DB
                            oDB.Data = sData
                            oDB.UpdateSQL(oclient.IsArbitrage, cmd)
                        End If
                    End If
                End If
                Application.DoEvents()
            End While
        End If
        cmd.CommandText = "exec ArbitragePriceAdjust "
        cmd.ExecuteScalar()
        conn.Close()
        RaiseEvent NewLine("")

        fs.Close()
        fs.Dispose()

        Return True
    End Function
#End Region 'Process CSV File

#Region " Process TSV File (Default) "
    Private Function Processfile_TSV(ByVal sFile As String, ByVal oclient As Clients, ByVal FileDate As Date, ByVal nCountLimit As Integer) As Boolean
        Dim nCount As Integer
        Dim sHeader As String = ""
        Dim bResult As Boolean = False

        Dim fs As StreamReader = File.OpenText(sFile)
        For nCount = 1 To nCountLimit
            sHeader = fs.ReadLine
        Next

        fs.Close()
        fs.Dispose()

        Dim sColumns() As String = sHeader.Split(vbTab)
        Dim bTactic As Boolean = False

        If sColumns.GetUpperBound(0) >= 2 Then
            For nCount = 0 To sColumns.GetUpperBound(0)
                If sColumns(2).ToUpper = "TACTIC ID" Then bTactic = True
            Next
        End If

        If bTactic Then
            Debug.Print("Ad Group File")
            bResult = ProcessAdGroupfile_TSV(sFile, oclient, FileDate, nCountLimit)
        Else
            Debug.Print("Keyword File")
            bResult = ProcessKeywordfile_TSV(sFile, oclient, FileDate, nCountLimit)
        End If

        Return bResult
    End Function

    Private Function ProcessKeywordfile_TSV(ByVal sFile As String, ByVal oclient As Clients, ByVal FileDate As Date, ByVal nCountLimit As Integer) As Boolean
        Debug.Print("Processing Keyword TSV")
        Dim nCntLimit As Integer = nCountLimit + 1

        Dim sYahooAcctID As String = oclient.Yahoo.AccountID
        Dim sYahooSubAcctID As String = oclient.Yahoo.SubAccountID

        Dim sCustomerID As String = oclient.CustID
        Dim sCustomerName As String = oclient.Name

        Dim sHeader(nCntLimit) As String
        Dim nCount As Integer
        Dim fs As StreamReader = File.OpenText(sFile)
        For nCount = 1 To nCntLimit
            sHeader(nCount) = fs.ReadLine
        Next

        Dim oDB As New clDatabase
        oDB.cs = ConnectionString
        oDB.EngineID = 1
        oDB.myDate = FileDate

        oDB.CustomerID = sCustomerID
        oDB.CustomerName = sCustomerName

        oDB.Columns = ParseKeywordColumns(sHeader(nCountLimit))
        oDB.DataTypes = _ColTypes

        oDB.FieldWithYahooAccountID = nFieldAcctID
        oDB.FieldWithYahooAccountName = nFieldAcctName
        oDB.YahooAccount = sYahooAcctID

        Dim nLine As Long = 0
        Dim sLine As String
        Dim conn As New SqlConnection(ConnectionString)
        Dim cmd As New SqlCommand
        conn.Open()
        cmd.Connection = conn
        If sYahooAcctID = "214045" Then
            'Yahoo Multi Account
            oDB.MemberOfMaster = True
            While Not fs.EndOfStream
                sLine = fs.ReadLine
                nLine += 1
                RaiseEvent NewLine(nLine.ToString)

                oDB.Data = sLine.Split(vbTab)
                If oDB.Data.GetUpperBound(0) = oDB.Columns.GetUpperBound(0) Then
                    oDB.YahooAccount = oDB.Data(nFieldAcctID) 'Set for each line
                    oDB.UpdateSQL(oclient.IsArbitrage, cmd)
                End If
                Application.DoEvents()
            End While
        Else
            'Yahoo Single Account
            While Not fs.EndOfStream
                sLine = fs.ReadLine
                nLine += 1
                RaiseEvent NewLine(nLine)

                oDB.Data = sLine.Split(vbTab)
                oDB.UpdateSQL(oclient.IsArbitrage, cmd)
                Application.DoEvents()
            End While
        End If
        cmd.CommandText = "exec ArbitragePriceAdjust "
        cmd.ExecuteScalar()

        conn.Close()
        RaiseEvent NewLine("")

        fs.Close()
        fs.Dispose()

        Return True
    End Function

    Private Function ProcessAdGroupfile_TSV(ByVal sFile As String, ByVal oclient As Clients, ByVal FileDate As Date, ByVal nCountLimit As Integer) As Boolean
        Debug.Print("Processing AdGroup TSV")
        Dim nCntLimit As Integer = nCountLimit + 1

        Dim sYahooAcctID As String = oclient.Yahoo.AccountID
        Dim sYahooSubAcctID As String = oclient.Yahoo.SubAccountID

        Dim sCustomerID As String = oclient.CustID
        Dim sCustomerName As String = oclient.Name

        Dim sHeader(nCntLimit) As String
        Dim nCount As Integer
        Dim fs As StreamReader = File.OpenText(sFile)
        For nCount = 1 To nCntLimit
            sHeader(nCount) = fs.ReadLine
        Next

        Dim sData() As String = Nothing
        Dim Index As Integer = 0

        Dim oDB As New clDatabase
        oDB.cs = ConnectionString
        oDB.EngineID = 1
        oDB.myDate = FileDate

        oDB.CustomerID = sCustomerID
        oDB.CustomerName = sCustomerName

        oDB.Columns = ParseAdGroupColumns(sHeader(nCountLimit))
        oDB.DataTypes = _ColTypes

        oDB.FieldWithYahooAccountID = nFieldAcctID
        oDB.FieldWithYahooAccountName = nFieldAcctName
        oDB.YahooAccount = sYahooAcctID

        Dim nLine As Long = 0
        Dim sLine As String
        Dim conn As New SqlConnection(ConnectionString)
        Dim cmd As New SqlCommand
        conn.Open()
        cmd.Connection = conn
        If sYahooAcctID = "214045" Then
            'Yahoo Multi Account
            oDB.MemberOfMaster = True
            While Not fs.EndOfStream
                sLine = fs.ReadLine
                nLine += 1
                RaiseEvent NewLine(nLine.ToString)

                sData = sLine.Split(vbTab)
                Index = sData.GetUpperBound(0) + 1
                ReDim Preserve sData(Index)
                sData(Index) = "Content Match"

                If sData.GetUpperBound(0) = oDB.Columns.GetUpperBound(0) Then
                    If nFieldTacticID >= 0 Then
                        If sData(nFieldTacticID) = IDS_ContentMatchID Then
                            'Updata DB
                            oDB.YahooAccount = sData(nFieldAcctID) 'Set for each line
                            oDB.Data = sData
                            oDB.UpdateSQL(oclient.IsArbitrage, cmd)
                        End If
                    End If
                End If
                Application.DoEvents()
            End While
        Else
            'Yahoo Single Account
            While Not fs.EndOfStream
                sLine = fs.ReadLine
                nLine += 1
                RaiseEvent NewLine(nLine)

                sData = sLine.Split(vbTab)
                Index = sData.GetUpperBound(0) + 1
                ReDim Preserve sData(Index)
                sData(Index) = "Content Match"

                If sData.GetUpperBound(0) = oDB.Columns.GetUpperBound(0) Then
                    If nFieldTacticID >= 0 Then
                        If sData(nFieldTacticID) = IDS_ContentMatchID Then
                            'Updata DB
                            oDB.Data = sData
                            oDB.UpdateSQL(oclient.IsArbitrage,cmd)
                        End If
                    End If
                End If
                Application.DoEvents()
            End While
        End If
        cmd.CommandText = "exec ArbitragePriceAdjust "
        cmd.ExecuteScalar()
        conn.Close()
        RaiseEvent NewLine("")

        fs.Close()
        fs.Dispose()

        Return True
    End Function
#End Region 'Process TSV File (Default)

#Region " Process XML File "
    Private Function Processfile_XML(ByVal sFile As String, ByVal oclient As Clients, ByVal FileDate As Date) As Boolean
        Debug.Print("Processing XML")

        '****************************************************
        ' Still Needs Work !!!
        '****************************************************

        Dim reader As New XmlTextReader(sFile)
        reader.WhitespaceHandling = WhitespaceHandling.None

        '<row>
        Dim scmpgnID As String
        Dim sadGrpID As String
        Dim skeywordID As String
        Dim surlID As String
        Dim saccountName As String
        Dim saccountID As String
        Dim scmpgnName As String
        Dim sadGrpName As String
        Dim skeywordName As String
        Dim surl As String

        '<analytics>
        Dim snumImpr As String
        Dim snumClick As String
        Dim sctr As String
        Dim scpc As String
        Dim snumConv As String
        Dim sclickConvRate As String
        Dim srevenue As String
        Dim sroas As String
        Dim scost As String
        Dim saveragePosition As String

        '</analytics>
        '</row>
        While reader.Read()
            If reader.Name.ToLower = "row" Then
                scmpgnID = reader.GetAttribute("cmpgnID")
                sadGrpID = reader.GetAttribute("adGrpID")
                skeywordID = reader.GetAttribute("keywordID")
                surlID = reader.GetAttribute("urlID")
                saccountName = reader.GetAttribute("accountName")
                saccountID = reader.GetAttribute("accountID")
                scmpgnName = reader.GetAttribute("cmpgnName")
                sadGrpName = reader.GetAttribute("adGrpName")
                skeywordName = reader.GetAttribute("keywordName")
                surl = reader.GetAttribute("url")

                'Need to read analytic line
                snumImpr = reader.GetAttribute("numImpr")
                snumClick = reader.GetAttribute("numClick")
                sctr = reader.GetAttribute("ctr")
                scpc = reader.GetAttribute("cpc")
                snumConv = reader.GetAttribute("numConv")
                sclickConvRate = reader.GetAttribute("clickConvRate")
                srevenue = reader.GetAttribute("revenue")
                sroas = reader.GetAttribute("roas")
                scost = reader.GetAttribute("cost")
                saveragePosition = reader.GetAttribute("averagePosition")

                'Debug.Print("Line: " & reader.LineNumber & " Date: " & sDate & " Campaign: " & sCampaign & " Keywords: " & sKwSite & " Imps: " & sImps)
            End If
        End While
        reader.Close()


        Return True
    End Function
#End Region 'Process XML File

#Region " Private Methods "
    Private Function ProcessFile(ByVal oFile As FileInfo, ByVal oclient As Clients, ByVal FileDate As Date) As Boolean
        Dim nCount As Integer = 0
        Select Case oFile.Extension.ToUpper
            Case ".CSV"
                Return Processfile_Check(oFile.FullName, oclient, FileDate)

            Case ".TSV"
                Return Processfile_TSV(oFile.FullName, oclient, FileDate, nCount)

            Case ".XML"
                Return Processfile_XML(oFile.FullName, oclient, FileDate)

            Case Else
                MsgBox("Don't Know what to do with a " & oFile.Extension & " file.")

        End Select

        Return False
    End Function

    Private Function Processfile_Check(ByVal sfile As String, ByVal oclient As Clients, ByVal FileDate As Date) As Boolean
        'Check if Tab or Comma Seperated and Run correct process

        Dim sLine As String = ""
        Dim fs As StreamReader = File.OpenText(sfile)
        Dim ssCSV() As String
        Dim ssTSV() As String

        'Read past the headers
        Dim nCount, nSveCnt As Integer
        For nCount = 1 To 9
            sLine = fs.ReadLine

            ssCSV = sLine.Split(",")
            ssTSV = sLine.Split(vbTab)
            'Find Column Headers
            If ssCSV(0).StartsWith("Campaign Id") Or ssTSV(0).StartsWith("Campaign Id") Then
                nSveCnt = nCount
                nCount = 9
            End If
            'Validate complete data, if not do not process.
            If ssCSV(0).StartsWith("Warning:") Or ssTSV(0).StartsWith("Warning:") Then
                Debug.Print("Incomplete Yahoo Data, must rerun")
                fs.Close()
                fs.Dispose()
                Return False
            End If
        Next

        'Validate complete data, if not do not process.
        If sLine.Length = 0 Then
            Debug.Print("Incomplete Yahoo Data, must rerun")
            fs.Close()
            fs.Dispose()
            Return False
        End If

        fs.Close()
        fs.Dispose()
        Dim sCSV() As String = sLine.Split(",")
        Dim sTSV() As String = sLine.Split(vbTab)

        If sCSV.GetUpperBound(0) > sTSV.GetUpperBound(0) Then
            Return Processfile_CSV(sfile, oclient, FileDate, nSveCnt)
        Else
            Return Processfile_TSV(sfile, oclient, FileDate, nSveCnt)
        End If
    End Function

    Private Function ParseKeywordColumns(ByVal sLine As String) As String()
        Dim sColumns() As String = sLine.Split(vbTab)

        Dim nCount As Integer
        For nCount = sColumns.GetLowerBound(0) To sColumns.GetUpperBound(0)
            If sColumns(nCount).ToUpper = "ACCOUNT ID" Then nFieldAcctID = nCount
            If sColumns(nCount).ToUpper = "ACCOUNT" Then nFieldAcctName = nCount

            sColumns(nCount) = _ColumnsKeyword(sColumns(nCount))
        Next

        Return sColumns
    End Function

    Private Function ParseAdGroupColumns(ByVal sLine As String) As String()
        Dim sColumns() As String = sLine.Split(vbTab)

        'Add a Keyword Column since it is not there
        Dim Index As Integer = sColumns.GetUpperBound(0) + 1
        ReDim Preserve sColumns(Index)
        sColumns(Index) = "Keyword"

        nFieldTacticID = -1
        Dim nCount As Integer
        For nCount = sColumns.GetLowerBound(0) To sColumns.GetUpperBound(0)
            If sColumns(nCount).ToUpper = "ACCOUNT ID" Then nFieldAcctID = nCount
            If sColumns(nCount).ToUpper = "ACCOUNT" Then nFieldAcctName = nCount
            If sColumns(nCount).ToUpper = "TACTIC ID" Then nFieldTacticID = nCount

            sColumns(nCount) = _ColumnsAdGroup(sColumns(nCount))
        Next

        Return sColumns
    End Function

    Private Function ParseKeywordColumnsCSV(ByVal sLine As String) As String()
        If sLine Is Nothing Then
            Return Nothing
        Else

            Dim sColumns() As String = sLine.Split(",")

            Dim nCount As Integer
            For nCount = sColumns.GetLowerBound(0) To sColumns.GetUpperBound(0)
                If sColumns(nCount).ToUpper = "ACCOUNT ID" Then nFieldAcctID = nCount
                If sColumns(nCount).ToUpper = "ACCOUNT" Then nFieldAcctName = nCount

                sColumns(nCount) = _ColumnsKeyword(sColumns(nCount))
            Next

            Return sColumns
        End If
    End Function

    Private Function ParseAdGroupColumnsCSV(ByVal sLine As String) As String()
        Dim sColumns() As String = sLine.Split(",")

        'Add a Keyword Column since it is not there
        Dim Index As Integer = sColumns.GetUpperBound(0) + 1
        ReDim Preserve sColumns(Index)
        sColumns(Index) = "Keyword"

        nFieldTacticID = -1
        Dim nCount As Integer
        For nCount = sColumns.GetLowerBound(0) To sColumns.GetUpperBound(0)
            If sColumns(nCount).ToUpper = "ACCOUNT ID" Then nFieldAcctID = nCount
            If sColumns(nCount).ToUpper = "ACCOUNT" Then nFieldAcctName = nCount
            If sColumns(nCount).ToUpper = "TACTIC ID" Then nFieldTacticID = nCount

            sColumns(nCount) = _ColumnsAdGroup(sColumns(nCount))
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
    Public Sub ResetFilesEOM(ByVal Client As Clients, ByVal StartDate As Date, ByVal EndDate As Date)
        Dim sFileName As String = ""
        Dim MyFile As String = ""

        Dim sDir As String = GetMyDir(Client)
        Dim sDirP As String = sDir & "Processed\"

        Dim oDir As New DirectoryInfo(sDirP)
        Dim oFile As FileInfo

        Dim FileDate As Date

        If Directory.Exists(sDir) = False Then Directory.CreateDirectory(sDir)
        If Directory.Exists(sDirP) = False Then Directory.CreateDirectory(sDirP)

        If Client.Yahoo.MultiDayReport Then
            MyFile = "YM_" & Replace(Client.Name, " ", "") & Client.Yahoo.EndDate.ToString("_yyyy_MM_dd") & ".xml"
        Else
            MyFile = "YD_" & Replace(Client.Name, " ", "") & Client.Yahoo.EndDate.ToString("_yyyy_MM_dd") & ".xml"
        End If

        For Each oFile In oDir.GetFiles
            'FileDate = GetFileDate(oFile.Name)

            'Reset Nov EOM
            'If (oFile.Name.IndexOf("GM_") <> -1) And (oFile.Name.IndexOf("2009_11_30.xml") <> -1) Then
            '    'If oFile.Name.ToString = MyFile Then

            '    Dim sFile As String = sDir & oFile.Name
            '    If File.Exists(sFile) Then File.Delete(sFile)
            '    oFile.MoveTo(sFile)
            '    Application.DoEvents()

            'End If

            'December Reset
            'If (oFile.Name.IndexOf("YD") <> -1) And (oFile.Name.IndexOf("2009_12") <> -1) Then
            '    'If oFile.Name.ToString = MyFile Then

            '    Dim sFile As String = sDir & oFile.Name
            '    If File.Exists(sFile) Then File.Delete(sFile)
            '    oFile.MoveTo(sFile)
            '    Application.DoEvents()

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
        sBad(0) = "(Yahoo) Date: " & CheckDate.ToString("yyyy_MM_dd")
        Dim Index As Integer = 0

        For Each client As Clients In Clients
            If client.Yahoo.Has_AccountID Then
                Dim sDir As String = GetMyDir(client)
                If bCheckProcessedDir = True Then sDir = sDir & "Processed\"
                Dim bExists As Boolean = False

                Dim sFileName1 As String = ""
                Dim sFileName2 As String = ""
                Dim sFileName3 As String = ""

                'Key
                sFileName1 = sDir & "Y_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".csv"
                sFileName2 = sDir & "YD_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".csv"
                sFileName3 = sDir & "YM_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".csv"

                bExists = False
                If File.Exists(sFileName1) Then bExists = True
                If File.Exists(sFileName2) Then bExists = True
                If File.Exists(sFileName3) Then bExists = True

                If bExists = False Then
                    Index = sBad.GetUpperBound(0) + 1
                    ReDim Preserve sBad(Index)
                    sBad(Index) = "YAHOO(KEY) MISSING FOR (" & client.CustID.ToString & ") " & client.Name
                End If

                'CM
                sFileName1 = sDir & "Ya_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".csv"
                sFileName2 = sDir & "YDa_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".csv"
                sFileName3 = sDir & "YMa_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".csv"

                bExists = False
                If File.Exists(sFileName1) Then bExists = True
                If File.Exists(sFileName2) Then bExists = True
                If File.Exists(sFileName3) Then bExists = True

                If bExists = False Then
                    Index = sBad.GetUpperBound(0) + 1
                    ReDim Preserve sBad(Index)
                    sBad(Index) = "YAHOO(CM) MISSING FOR (" & client.CustID.ToString & ") " & client.Name
                End If

            End If
        Next

        If Index = 0 Then sBad(0) = ""
        Return sBad
    End Function
#End Region 'CheckProcessed

End Class
