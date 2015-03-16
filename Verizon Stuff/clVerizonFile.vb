Imports System.IO
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic

Public Class clVerizonFile
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
        Files = oSettings.SaveBaseLocation & "Verizon\"
        ConnectionString = oSettings.ConnectionString
        DirectoryNamesAsRecID = oSettings.DirectoryNamesAsRecID

        Dim nCount As Integer

        Dim Index As Integer
        For nCount = oSettings.VerizonColumns.GetLowerBound(0) To oSettings.VerizonColumns.GetUpperBound(0)
            If Not IsNothing(oSettings.VerizonColumns(nCount)) Then
                If IsNothing(sColumnNames) Then
                    Index = 0
                    ReDim sColumnNames(Index)
                    ReDim sColumns(Index)
                Else
                    Index = sColumnNames.GetUpperBound(0) + 1
                    ReDim Preserve sColumnNames(Index)
                    ReDim Preserve sColumns(Index)
                End If
                sColumnNames(Index) = oSettings.VerizonColumns(nCount).Name
                sColumns(Index) = oSettings.VerizonColumns(nCount).Value
                If sColumnNames(Index).ToLower = "date" Then
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
                MoveErrorFile(oFile)
            End If

            Application.DoEvents()
        Next

    End Sub
#End Region 'Public Sub Execute

#Region " Process XML File (Default)"
    Private Function Processfile_XML(ByVal sFile As String, ByVal oClient As Clients, ByVal FileDate As Date) As Boolean
        Debug.Print("Processing XML")

        Dim sVerizonAcctID As String = oClient.Verizon.AccountID

        Dim sCustomerID As String = oClient.CustID
        Dim sCustomerName As String = oClient.Name

        Dim oDB As New clDatabase
        oDB.cs = ConnectionString
        oDB.EngineID = 4
        oDB.myDate = FileDate
        oDB.CustomerID = sCustomerID
        oDB.CustomerName = sCustomerName

        oDB.Columns = sColumns
        oDB.DataTypes = _ColTypes

        Dim nCount As Integer

        Dim reader As New XmlTextReader(sFile)
        reader.WhitespaceHandling = WhitespaceHandling.None

        Dim nLine As Long = 0
        Dim sData() As String = Nothing
        Dim bData As Boolean = False
        Dim conn As New SqlConnection(ConnectionString)
        Dim cmd As New SqlCommand
        conn.Open()
        cmd.Connection = conn
        Try
            While reader.Read()
                Application.DoEvents()

                Select Case reader.Name
                    Case "row"
                        If reader.NodeType = XmlNodeType.Element Then
                            'Start of Row
                            bData = True

                            nLine += 1
                            RaiseEvent NewLine(nLine.ToString)

                            ReDim sData(sColumnNames.GetUpperBound(0))
                            For nCount = 0 To sColumnNames.GetUpperBound(0)
                                If sColumnNames(nCount) = "campaign" Then
                                    sData(nCount) = oClient.Name
                                    Exit For
                                End If
                            Next
                        Else
                            bData = False
                            oDB.Data = sData
                            oDB.myDate = FileDate

                            oDB.UpdateSQL(oClient.IsArbitrage, cmd)
                            Application.DoEvents()
                        End If

                    Case Else
                        'Row Data
                        If bData = True Then
                            For nCount = 0 To sColumnNames.GetUpperBound(0)
                                If reader.Name = sColumnNames(nCount) Then
                                    If reader.IsStartElement Then
                                        reader.Read()   'Get the Value
                                        sData(nCount) = reader.Value
                                        reader.Read()   'Get the end Element
                                    End If
                                End If
                            Next
                        End If
                End Select
            End While
            cmd.CommandText = "exec ArbitragePriceAdjust "
            cmd.ExecuteScalar()

            conn.Close()
            RaiseEvent NewLine("")

            reader.Close()

            Return True
        Catch e As XmlException  ' Report does not have a root node.
            reader.Close()
            Log(oClient.LogName(True) & "File: " & sFile & " error: " & e.Message)
            oClient.Verizon.ErrorReport = True
            Return False
        End Try
    End Function
#End Region 'Process XML File (Default)

#Region " Private Methods "
    Private Function ProcessFile(ByVal oFile As FileInfo, ByVal oClient As Clients, ByVal FileDate As Date) As Boolean
        Select Case oFile.Extension.ToUpper
            Case ".XML"
                Return Processfile_XML(oFile.FullName, oClient, FileDate)

            Case Else
                MsgBox("Don't Know what to do with a " & oFile.Extension & " file.")

        End Select

        Return False
    End Function

    Private Sub MoveFile(ByVal oFile As FileInfo)
        Dim sDir As String = oFile.DirectoryName & "\Processed\"
        Dim sFile As String = sDir & oFile.Name

        If Directory.Exists(sDir) = False Then Directory.CreateDirectory(sDir)
        If File.Exists(sFile) Then File.Delete(sFile)
        oFile.MoveTo(sFile)
    End Sub
    Private Sub MoveErrorFile(ByVal oFile As FileInfo)
        Dim sDir As String = oFile.DirectoryName & "\Error\"
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
        sBad(0) = "(Verizon) Date: " & CheckDate.ToString("yyyy_MM_dd")
        Dim Index As Integer = 0

        For Each client As Clients In Clients
            If client.Verizon.Has_AccountID Then
                Dim sDir As String = GetMyDir(client)
                If bCheckProcessedDir = True Then sDir = sDir & "Processed\"

                Dim sFileName1 As String = sDir & "V_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".xml"
                Dim sFileName2 As String = sDir & "VD_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".xml"
                Dim sFileName3 As String = sDir & "VM_" & Replace(client.Name, " ", "") & CheckDate.ToString("_yyyy_MM_dd") & ".xml"

                Dim bExists As Boolean = False
                If File.Exists(sFileName1) Then bExists = True
                If File.Exists(sFileName2) Then bExists = True
                If File.Exists(sFileName3) Then bExists = True

                If bExists = False Then
                    Index = sBad.GetUpperBound(0) + 1
                    ReDim Preserve sBad(Index)
                    sBad(Index) = "VERIZON MISSING FOR (" & client.CustID.ToString & ") " & client.Name
                End If
            End If
        Next

        If Index = 0 Then sBad(0) = ""
        Return sBad
    End Function
#End Region 'CheckProcessed

End Class
