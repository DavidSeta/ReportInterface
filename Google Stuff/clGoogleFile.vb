Imports System.IO
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic

Public Class clGoogleFile
    Private Files As String
    Private DirectoryNamesAsRecID As Boolean = False

    Private nColumnDate As Integer = -1
    Private sColumnNames() As String
    Private sColumns() As String
    Private _ColTypes As Hashtable = New Hashtable()
    Private ConnectionString As String

    Public Event NewLine(ByVal sLine As String)

#Region " Sub New "
    Sub New(ByVal oSettings As clSettings)
        Files = oSettings.SaveBaseLocation & "Google\"
        ConnectionString = oSettings.ConnectionString
        DirectoryNamesAsRecID = oSettings.DirectoryNamesAsRecID

        Dim nCount As Integer

        Dim Index As Integer
        For nCount = oSettings.GoogleColumns.GetLowerBound(0) To oSettings.GoogleColumns.GetUpperBound(0)
            If Not IsNothing(oSettings.GoogleColumns(nCount)) Then
                If IsNothing(sColumnNames) Then
                    Index = 0
                    ReDim sColumnNames(Index)
                    ReDim sColumns(Index)
                Else
                    Index = sColumnNames.GetUpperBound(0) + 1
                    ReDim Preserve sColumnNames(Index)
                    ReDim Preserve sColumns(Index)
                End If
                sColumnNames(Index) = oSettings.GoogleColumns(nCount).Name
                sColumns(Index) = oSettings.GoogleColumns(nCount).Value
                'V13 Patch = date
                If sColumnNames(Index).ToLower = "day" Then
                    Me.nColumnDate = Index
                    sColumns(Index) = ""
                End If
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
            If (oFile.Length = 0) Then
                Log(oClient.LogName(True) & "Empty File: " & oFile.Name)
                MoveFile(oFile)
            Else
                Log(oClient.LogName(True) & "File: " & oFile.Name)
                If ProcessFile(oFile, oClient, FileDate, bFlexEOM) Then MoveFile(oFile)
                Application.DoEvents()
            End If
        Next

    End Sub
#End Region 'Public Sub Execute

#Region " Public Sub Execute EOM"
    Public Function ExecuteEOM(ByVal oClient As Clients) As String
        Dim sRowDeleted As String
        Dim sGoogleAcctID As String = oClient.Google.AccountID
        Dim sCustomerID As String = oClient.CustID
        Dim sCustomerName As String = oClient.Name

        'Instantiate and set fields for use in SQL DELETE FROM PERFDATA
        Dim oDB As New clDatabase
        oDB.cs = ConnectionString
        oDB.EngineID = 2
        oDB.CustomerID = sCustomerID
        oDB.CustomerName = sCustomerName
        oDB.sEOMDate = oClient.Google.StartDate
        oDB.eEOMDate = oClient.Google.EndDate
        sRowDeleted = oDB.DeletePerfdata()

        Return sRowDeleted
    End Function
#End Region 'Public Sub Execute EOM

#Region " Process CSV File "
    Private Function Processfile_CSV(ByVal sFile As String, ByVal oClient As Clients, ByVal FileDate As Date, ByVal bFlexEOM As Boolean) As Boolean
        Debug.Print("Processing CSV")

        Return True
    End Function
#End Region 'Process CSV File

#Region " Process TSV File "
    Private Function Processfile_TSV(ByVal sFile As String, ByVal oClient As Clients, ByVal FileDate As Date, ByVal bFlexEOM As Boolean) As Boolean
        Debug.Print("Processing TSV")

        Return True
    End Function
#End Region 'Process TSV File

#Region " Process XML File (Default)"
    Private Function Processfile_XML(ByVal sFile As String, ByVal oClient As Clients, ByVal FileDate As Date, ByVal bFlexEOM As Boolean) As Boolean
        Debug.Print("Processing XML")

        Dim sGoogleAcctID As String = oClient.Google.AccountID
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
        cmd.CommandTimeout = 0
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
                        'If Not IsNothing(reader.GetAttribute(sColumnNames(nCount))) Then
                        bHaveData = True
                        'Move value to array
                        sData(nCount) = reader.GetAttribute(sColumnNames(nCount))

                        'Feb 1 2010 will need to divide by 1000000 can force via setting header value in 
                        ' GoogleReport.vb>SaveFile()   client.Headers.Add("returnMoneyInMicros: true")
                        ''V201003 No longer require Money treatment and consideration for "kwsite".
                        ''Money treatment - Google uses 1000000 multiplier
                        If sColumnNames(nCount).ToLower = "cost" Then
                            Dim nTemp As Double = reader.GetAttribute(sColumnNames(nCount))
                            nTemp = nTemp / 1000000
                            sData(nCount) = nTemp.ToString
                        End If
                        If sColumnNames(nCount).ToLower = "avgcpc" Then
                            Dim nTemp As Double = reader.GetAttribute(sColumnNames(nCount))
                            nTemp = nTemp / 1000000
                            sData(nCount) = nTemp.ToString
                        End If
                        If sColumnNames(nCount).ToLower = "cpc" Then
                            Dim nTemp As Double = reader.GetAttribute(sColumnNames(nCount))
                            nTemp = nTemp / 1000000
                            sData(nCount) = nTemp.ToString
                        End If
                        If sColumnNames(nCount).ToLower = "costconv1perclick" Then
                            Dim nTemp As Double = reader.GetAttribute(sColumnNames(nCount))
                            nTemp = nTemp / 1000000
                            sData(nCount) = nTemp.ToString
                        End If
                        If sColumnNames(nCount).ToLower = "costperconv" Then
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

                    'No Commas in numeric data
                    If IsNumeric(sData(nCount)) Then
                        sData(nCount) = Replace(sData(nCount), ",", "")
                    End If

                Next    'For nCount


                 
                'Must have data to proceed.
                If ((bHaveData) And (sData(1) <> "0")) Then
                    oDB.Data = sData
                    If Not IsNothing(oDB.Data(nColumnDate)) Then
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
                End If 'If we had any data to insert/update into our DB

            End If   'If reader.Name.ToLower = "row"
        End While    'While reader.Read()

        ''Stored Procedure to populate PerfDataArbitrage.ModifiedCost based on Customers.ArbitrageMarkUP
        'cmd.CommandText = "exec ArbitragePriceAdjust "
        'cmd.CommandTimeout = 0
        'cmd.ExecuteScalar()

        conn.Close()

        RaiseEvent NewLine("")
        reader.Close()
        Return True
    End Function
#End Region 'Process XML File (Default)

#Region " Private Methods "
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
            Return Processfile_CSV(sfile, oClient, FileDate, bFlexEOM)
        Else
            Return Processfile_TSV(sfile, oClient, FileDate, bFlexEOM)
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
    Public Function UpdateRemarketing_Keywords(ByVal psDate As Date, ByVal peDate As Date) As String
        Debug.Print("Update Google Remarketing Keyword")
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

        If Client.Google.MultiDayReport Then
            MyFile = "GM_" & Replace(Client.Name, " ", "") & Client.Google.EndDate.ToString("_yyyy_MM_dd") & ".xml"
        Else
            MyFile = "GD_" & Replace(Client.Name, " ", "") & Client.Google.EndDate.ToString("_yyyy_MM_dd") & ".xml"
        End If

        For Each oFile In oDir.GetFiles
            'FileDate = GetFileDate(oFile.Name)
            If oFile.Name.ToString = MyFile Then

                Dim sFile As String = sDir & oFile.Name
                If File.Exists(sFile) Then File.Delete(sFile)
                oFile.MoveTo(sFile)
                Application.DoEvents()
            End If
            'Reset Feb EOM
            'If (oFile.Name.IndexOf("GM_") <> -1) And (oFile.Name.IndexOf("2011_02_28.xml") <> -1) Then
            '    If oFile.Name.ToString = MyFile Then

            '        Dim sFile As String = sDir & oFile.Name
            '        If File.Exists(sFile) Then File.Delete(sFile)
            '        oFile.MoveTo(sFile)
            '        Application.DoEvents()
            '    End If
            'End If

            ''January Reset
            'If (oFile.Name.IndexOf("2010") <> -1) Then
            '    'If oFile.Name.ToString = MyFile Then

            '    Dim sFile As String = sDir & oFile.Name
            '    If File.Exists(sFile) Then File.Delete(sFile)
            '    oFile.MoveTo(sFile)
            '    Application.DoEvents()

            'End If
            Application.DoEvents()
        Next
    End Sub
#End Region 'Reset DB Files Sub

#Region " CheckProcessed "
    Public Function CheckNotProcessed(ByVal Clients() As Clients, ByVal CheckDate As Date, ByVal bCheckProcessedDir As Boolean) As String()
        Dim sBad(0) As String
        sBad(0) = "(Google) Date: " & CheckDate.ToString("yyyy_MM_dd")
        Dim Index As Integer = 0

        For Each client As Clients In Clients
            If client.Google.Has_AccountID Then

                Dim sDir As String = GetMyDir(client)
                If bCheckProcessedDir = True Then sDir = sDir & "Processed\"

                Dim sFileName1 As String = sDir & "G_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".xml"
                Dim sFileName2 As String = sDir & "GD_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".xml"
                Dim sFileName3 As String = sDir & "GM_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".xml"

                Dim bExists As Boolean = False
                If File.Exists(sFileName1) Then bExists = True
                If File.Exists(sFileName2) Then bExists = True
                If File.Exists(sFileName3) Then bExists = True

                If bExists = False Then
                    Index = sBad.GetUpperBound(0) + 1
                    ReDim Preserve sBad(Index)
                    sBad(Index) = "GOOGLE MISSING FOR (" & client.CustID.ToString & ") " & client.Name
                End If
            End If
        Next

        If Index = 0 Then sBad(0) = ""
        Return sBad
    End Function
#End Region 'CheckProcessed
End Class
