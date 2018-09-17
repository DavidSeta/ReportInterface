Imports System.Data.SqlClient

Public Class clDatabase
    Public cs As String
    Private conn As SqlConnection

    Private _MemberOfMaster As Integer = 0

    'Private _DataTypes() As String
    Public DataTypes As Hashtable

    Private AccountID As Integer = 0
    Public sEOMDate As Date
    Public eEOMDate As Date

    Public myDate As Date
    Public EngineID As Integer = 0
    Public Columns() As String
    Public Data() As String

    Public YahooAccount As String = ""
    Public YahooAccountLast As String = ""
    Public YahooAccountActive As Boolean = True
    Public FieldWithYahooAccountName As Integer = 0
    Public FieldWithYahooAccountID As Integer = 0

    Private myHash As Hashtable = New Hashtable()

    Public CustomerID As String = ""
    Public CustomerName As String = ""

    Private Function MOM() As String
        Return _MemberOfMaster.ToString
    End Function

#Region " Public Properties & Such "

    Public Function IsdataValid() As Boolean
        Return IIf(Columns.GetUpperBound(0) = Data.GetUpperBound(0), True, False)
    End Function

    Public Property MemberOfMaster() As Boolean
        Get
            Return IIf(_MemberOfMaster = 1, True, False)
        End Get
        Set(ByVal value As Boolean)
            _MemberOfMaster = IIf(value = True, 1, 0)
        End Set
    End Property

    Private Sub SetAccountID()
        Dim cmd As SqlCommand
        Dim Status As Integer

        Dim SQL As String
        SQL = "SELECT CustomerID, Status " & _
              "FROM X_Client_Campaign " & _
              "WHERE  EngineID = " & Me.EngineID.ToString & " " & _
              "AND AccountID = '" & YahooAccount & "'"

        Dim conn As New SqlConnection(cs)
        conn.Open()

        cmd = New SqlCommand(SQL, conn)
        Dim rdr As SqlDataReader = cmd.ExecuteReader

        AccountID = 0
        Status = 1
        If rdr.Read Then
            AccountID = rdr("CustomerID")
            If Not IsDBNull(rdr("Status")) Then Status = rdr("Status")
        End If
        YahooAccountActive = IIf(Status = 1, True, False)

        If AccountID = 0 Then
            'Insert New Account
            InsertNewAccount(conn)

            'Get new ID
            cmd = New SqlCommand(SQL, conn)
            AccountID = cmd.ExecuteScalar
        End If
        CustomerID = AccountID.ToString

        cmd.Dispose()
        conn.Close()
        conn.Dispose()
    End Sub

    Private Sub InsertNewAccount(ByRef conn As SqlConnection)
        Dim nCustID As Integer = GetCustomerIDFromCustomers(conn)

        Dim SQL As String
        If Me.MemberOfMaster Then
            SQL = "INSERT INTO X_Client_Campaign (CustomerID,EngineID, Status, MemberOfMaster, AccountID)" & _
                  "VALUES (" & nCustID.ToString & "," & Me.EngineID.ToString & ",1," & MOM() & "," & Data(Me.FieldWithYahooAccountID) & ")"
        Else
            SQL = "INSERT INTO X_Client_Campaign (CustomerID,EngineID, Status, MemberOfMaster, AccountID)" & _
                  "VALUES (" & nCustID.ToString & "," & Me.EngineID.ToString & ",1," & MOM() & "," & AccountID.ToString & ")"
        End If

        Dim cmd As New SqlCommand(SQL, conn)
        cmd.ExecuteScalar()
        cmd.Dispose()
    End Sub

    Private Function GetCustomerIDFromCustomers(ByRef conn As SqlConnection) As Integer
        Dim cmd As SqlCommand
        Dim nAcctID As Integer = 0

        Dim SQL As String
        SQL = "SELECT CustomerID " & _
              "FROM Customers " & _
              "WHERE CustomerName = '" & Replace(Data(Me.FieldWithYahooAccountName), "'", "''") & "'"

        cmd = New SqlCommand(SQL, conn)
        nAcctID = cmd.ExecuteScalar

        If nAcctID < 1 Then
            AddCustomerToCustomers(conn)

            cmd = New SqlCommand(SQL, conn)
            nAcctID = cmd.ExecuteScalar
        End If
        cmd.Dispose()

        Return nAcctID
    End Function

    Private Sub AddCustomerToCustomers(ByRef conn As SqlConnection)
        Dim SQL As String
        SQL = "INSERT INTO Customers (CustomerName, Active) " & _
              "VALUES ('" & Replace(Data(Me.FieldWithYahooAccountName), "'", "''") & "',0)"

        Dim cmd As New SqlCommand(SQL, conn)
        cmd.ExecuteScalar()
        cmd.Dispose()
    End Sub

#End Region 'Public Properties & Such

#Region " Delete Perfdata"
    '====================================================================================================
    'DeletePerfdata () Delete Perfdata for EOM processing                                                
    'Invoked from myRow() function used in UpdateSQL()
    '====================================================================================================
    Public Function DeletePerfdata() As String
        Dim ra As Integer

        Dim sWhere As String = "WHERE EngineID=" & EngineID.ToString & " " & _
              "AND CustomerID = " & CustomerID & " "

        sWhere = sWhere & "AND [Date] BETWEEN '" & sEOMDate.ToString("yyyyMMdd") & "' AND '" & eEOMDate.ToString("yyyyMMdd") & "'"

        Dim SQL As String
        SQL = "DELETE PerfData " & sWhere

        Dim conn As New SqlConnection(cs)
        conn.Open()

        Dim cmd As New SqlCommand(SQL, conn)
        cmd.CommandTimeout = 0
        ra = cmd.ExecuteNonQuery()


        cmd.Dispose()
        conn.Close()
        conn.Dispose()

        Return ra.ToString

    End Function

#End Region 'Delete Perfdata 

#Region " GetWhere Function & Such "
    '====================================================================================================
    'GetWhere () Returns the string for the row based on the Client/Engine/Date/Campaign/AdGroup/Keyword 
    'Invoked from myRow() function used in UpdateSQL()
    '====================================================================================================
    Private Function GetWhereGA() As String
        Dim sWhere As String = "WHERE EngineID=" & EngineID.ToString & " " & _
              "AND CustomerID = " & CustomerID & " "
               
        'Only add condition if dimension found
        Dim sTemp As String
        sTemp = GetDataForGA("Keyword")
        If (sTemp <> "") Then
            sWhere = sWhere & "AND Keyword = '" & sTemp & "' "
        End If

        sTemp = GetDataForGA("AdGroup")
        If (sTemp <> "") Then
            sWhere = sWhere & "AND AdGroup = '" & sTemp & "' "
        End If

        sTemp = GetDataForGA("Campaign")
        If (sTemp <> "") Then
            sWhere = sWhere & "AND Campaign = '" & sTemp & "' "
        End If

        sWhere = sWhere & "AND [Date] = '" & myDate.ToString("yyyyMMdd") & "'"

        Return sWhere
    End Function
    '====================================================================================================
    'GetWhereGAV() Returns the string for the row based on the Client/Engine/Date/Campaign/AdGroup/Keyword 
    'Invoked from myRow() function used in UpdateSQL()
    '====================================================================================================
    Private Function GetWhereGAV() As String
        Dim sWhere As String = "WHERE EngineID=" & EngineID.ToString & " " & _
              "AND CustomerID = " & CustomerID & " "

        'Only add condition if dimension found
        Dim sTemp As String
        sTemp = GetDataForGA("CustomVarName")
        If (sTemp <> "") Then
            sWhere = sWhere & "AND CustomVarName = '" & sTemp & "' "
        End If

        sTemp = GetDataForGA("CustomVarValue")
        If (sTemp <> "") Then
            sWhere = sWhere & "AND CustomVarValue = '" & sTemp & "' "
        End If

        sWhere = sWhere & "AND [Date] = '" & myDate.ToString("yyyyMMdd") & "'"

        Return sWhere
    End Function
    '====================================================================================================
    'GetDataFor(sColumn) Returns the string for the columns - Campaign/AdGroup/Keyword.
    'Invoked from GetWhere() which is invoked by myRow() function used in UpdateSQL()
    '====================================================================================================
    Private Function GetDataForGA(ByVal sColumn As String) As String
        If sColumn = "" Then Return ""

        Dim sResult As String = ""
        Dim nCount As Integer

        If Not IsNothing(Columns) Then
            For nCount = Columns.GetLowerBound(0) To Columns.GetUpperBound(0)
                If Not IsNothing(Columns(nCount)) Then
                    If Columns(nCount).ToUpper = sColumn.ToUpper Then
                        sResult = Replace(Data(nCount), "'", "''")
                        Exit For
                    End If
                End If
            Next

            'Dec 2014 -  Bing can return blanks.  Consolidate changed to accomodate.
            If Me.EngineID = 3 Then
                If (sResult = "(not set)") Then
                    sResult = ""
                End If
            End If
            'When value is not set - return blank
            'March 2014 - Problem with returning blank can manifest itself via overwriting the first records existing revenue
            'Letting system put (not set) into the database for Campaign|AdGroup|Keyword
            'If (sResult = "(not set)") Then
            '    sResult = ""
            'End If
        End If

        Return sResult
    End Function
    '====================================================================================================
    'GetWhere () Returns the string for the row based on the Client/Engine/Date/Campaign/AdGroup/Keyword 
    'Invoked from myRow() function used in UpdateSQL()
    '====================================================================================================
    Private Function GetWhere() As String
        Dim sWhere As String = "WHERE EngineID=" & EngineID.ToString & " " & _
              "AND CustomerID = " & CustomerID & " " & _
              "AND Keyword = '" & GetDataFor("Keyword") & "' " & _
              "AND AdGroup = '" & GetDataFor("AdGroup") & "' " & _
              "AND Campaign = '" & GetDataFor("Campaign") & "' " & _
              "AND [Date] = '" & myDate.ToString("yyyyMMdd") & "'"

        Return sWhere
    End Function
    '====================================================================================================
    'GetDataFor(sColumn) Returns the string for the columns - Campaign/AdGroup/Keyword.
    'Invoked from GetWhere() which is invoked by myRow() function used in UpdateSQL()
    '====================================================================================================
    Private Function GetDataFor(ByVal sColumn As String) As String
        If sColumn = "" Then Return ""

        Dim sResult As String = ""
        Dim nCount As Integer

        If Not IsNothing(Columns) Then
            For nCount = Columns.GetLowerBound(0) To Columns.GetUpperBound(0)
                If Not IsNothing(Columns(nCount)) Then
                    If Columns(nCount).ToUpper = sColumn.ToUpper Then
                        sResult = Replace(Data(nCount), "'", "''")
                        Exit For
                    End If
                End If
            Next
        End If

        Return sResult
    End Function
#End Region 'GetWhere Function & Such

#Region " UpdateSQL Function and Such "
    '=====================================================================================================
    'UpdateSQLGAV() - Perform API result processing for insert/update into TireChoiceGAData Data                
    '=====================================================================================================
    Public Function UpdateSQLGAV(ByVal isArbitrage As Integer, ByVal cmd As SqlCommand) As Boolean

        Dim bResult As Boolean = True
        'Skip if output/result from API call does not map to expected columns
        If Data.GetUpperBound(0) <> Columns.GetUpperBound(0) Then Return bResult

        'Sub routine to combine/consolidate result data before Perfdata update/add
        ConsolidateV()

        'Use myRowGA() to return row or zero designating update/add
        Dim nRow As Integer = myRowGAV(isArbitrage)
        If nRow > 0 Then
            If isArbitrage = 1 Then
                bResult = UpdateRowArbitrageV(nRow, cmd)
            Else
                bResult = UpdateRowRegularV(nRow, cmd)
            End If
        Else
            If isArbitrage = 1 Then
                bResult = AddRowArbitrageV(cmd)
            Else
                bResult = AddRowRegularV(cmd)
            End If
        End If

        Return bResult
    End Function
    '=====================================================================================================
    'UpdateSQLGA() - Perform API result processing for insert/update into PerfData GA Data       
    ' 3/10/2015 - switch from Consolidate to ConsolidateGA
    '=====================================================================================================
    Public Function UpdateSQLGA(ByVal isArbitrage As Integer, ByVal cmd As SqlCommand) As Boolean

        Dim bResult As Boolean = True
        'Skip if output/result from API call does not map to expected columns
        If Data.GetUpperBound(0) <> Columns.GetUpperBound(0) Then Return bResult

        'Sub routine to combine/consolidate result data before Perfdata update/add
        ConsolidateGA()

        'Use myRowGA() to return row or zero designating update/add
        Dim nRow As Integer = myRowGA(isArbitrage)
        If nRow > 0 Then
            If isArbitrage = 1 Then
                bResult = UpdateRowArbitrage(nRow, cmd)
            Else
                bResult = UpdateRowRegular(nRow, cmd)
            End If
        Else
            If isArbitrage = 1 Then
                bResult = AddRowArbitrage(cmd)
            Else
                bResult = AddRowRegular(cmd)
            End If
        End If

        Return bResult
    End Function
    '=====================================================================================================
    'UpdateSQLFlexEOM() - Perform API result processing for insert/update into PerfData                  
    '                     Only allow update to the conversion
    '=====================================================================================================
    Public Function UpdateSQLFlexEOM(ByVal isArbitrage As Integer, ByVal cmd As SqlCommand) As Boolean
         
        Dim bResult As Boolean = True
        'Skip if output/result from API call does not map to expected columns
        If Data.GetUpperBound(0) <> Columns.GetUpperBound(0) Then Return bResult

        'Sub routine to combine/consolidate result data before Perfdata update/add
        Consolidate()

        'Use myRow() to return row or zero designating update/add
        Dim nRow As Integer = myRowFlexEOM(isArbitrage, cmd)

        If nRow > 0 Then
            If isArbitrage = 1 Then
                bResult = UpdateRowArbitrageFlexEOM(nRow, cmd)
            End If
        End If

        Return bResult
    End Function
    '=====================================================================================================
    'UpdateSQL() - Perform API result processing for insert/update into PerfData                          
    '=====================================================================================================
    Public Function UpdateSQL(ByVal isArbitrage As Integer, ByVal cmd As SqlCommand) As Boolean
        'Yahoo logic
        If EngineID = 1 Then
            If YahooAccount <> YahooAccountLast Then SetAccountID()
            YahooAccountLast = YahooAccount
            If YahooAccountActive = False Then Return True
        End If

        Dim bResult As Boolean = True
        'Skip if output/result from API call does not map to expected columns
        If Data.GetUpperBound(0) <> Columns.GetUpperBound(0) Then Return bResult

        'Sub routine to combine/consolidate result data before Perfdata update/add
        Consolidate()

        'Use myRow() to return row or zero designating update/add
        Dim nRow As Integer = myRow(isArbitrage, cmd)

        If nRow > 0 Then
            If isArbitrage = 1 Then
                bResult = UpdateRowArbitrage(nRow, cmd)
            Else
                bResult = UpdateRowRegular(nRow, cmd)
            End If
        Else
            If isArbitrage = 1 Then
                bResult = AddRowArbitrage(cmd)
            Else
                bResult = AddRowRegular(cmd)
            End If
        End If

        Return bResult
    End Function
    '========================================================================================================
    'Consolidate() - Consolidate - calling ConsolidateColumn() and AverageColumn() for numeric data          
    '========================================================================================================
    Private Sub Consolidate()
        Dim ncount As Integer

        'Columns numbers for report
        Dim cCampaign As Integer = -1
        Dim cAdGroup As Integer = -1
        Dim cKeyword As Integer = -1
        Dim cImpressions As Integer = -1
        Dim cClicks As Integer = -1
        Dim cCTR As Integer = -1
        Dim cCPC As Integer = -1
        Dim cConversions As Integer = -1
        Dim cConvValue As Integer = -1
        Dim cViewThroughConv As Integer = -1
        Dim cConvOneClick As Integer = -1
        Dim cClickConvRate As Integer = -1
        Dim cCostPerConv As Integer = -1
        Dim cRevenue As Integer = -1
        Dim cROAS As Integer = -1
        Dim cCost As Integer = -1
        Dim cAssists As Integer = -1
        Dim cAvgPos As Integer = -1
        
        Dim cTrans As Integer = -1

        'Locate columns assigning integer to variables
        For ncount = 0 To Columns.GetUpperBound(0)
            If Not Columns(ncount) Is Nothing Then
                Select Case Columns(ncount).ToLower
                    Case "campaign" : cCampaign = ncount
                    Case "adgroup" : cAdGroup = ncount
                    Case "keyword" : cKeyword = ncount
                    Case "impressions" : cImpressions = ncount
                    Case "clicks" : cClicks = ncount
                    Case "ctr" : cCTR = ncount
                    Case "cpc" : cCPC = ncount
                    Case "conversions" : cConversions = ncount
                    Case "viewthroughconv" : cViewThroughConv = ncount
                    Case "convoneclick" : cConvOneClick = ncount
                    Case "convvalue" : cConvValue = ncount
                    Case "clickconvrate" : cClickConvRate = ncount
                    Case "costperconv" : cCostPerConv = ncount
                    Case "revenue" : cRevenue = ncount
                    Case "roas" : cROAS = ncount
                    Case "cost" : cCost = ncount
                    Case "assists" : cAssists = ncount
                    Case "avgpos" : cAvgPos = ncount
                    Case "trans" : cTrans = ncount
                End Select
            End If
        Next

        'No data get out
        If cCampaign < 0 Then Exit Sub
        If cAdGroup < 0 Then Exit Sub
        If cKeyword < 0 Then Exit Sub

        'Verbiage for string columns assigned
        Dim sCampaign As String = Data(cCampaign)
        Dim sAdGroup As String = Data(cAdGroup)
        Dim sKeyword As String = Data(cKeyword)
        Dim sHashPre As String
        'adCenter has (Not Set) for AdGroup value, so eliminating when consolidating
        ' GetWhereGA()->GetDataForGA("AdGroup") will return a ""  
        'If Me.EngineID = 3 Then
        '    sHashPre = myDate.ToString("yyyyMMdd") & "." & sCampaign & "." & sKeyword & "."
        'Else
        '    sHashPre = myDate.ToString("yyyyMMdd") & "." & sCampaign & "." & sAdGroup & "." & sKeyword & "."
        'End If

        sHashPre = myDate.ToString("yyyyMMdd") & "." & sCampaign & "." & sAdGroup & "." & sKeyword & "."

        'Perform column consolidation for numeric columns
        ConsolidateColumn(cImpressions, sHashPre & "Impressions")
        ConsolidateColumn(cClicks, sHashPre & "Clicks")
        ConsolidateColumn(cCTR, sHashPre & "CTR")
        ConsolidateColumn(cCPC, sHashPre & "CPC")
        ConsolidateColumn(cConversions, sHashPre & "Conversions")
        ConsolidateColumn(cClickConvRate, sHashPre & "ClickConversionRate")
        ConsolidateColumn(cCostPerConv, sHashPre & "CostPerConv")
        ConsolidateColumn(cRevenue, sHashPre & "Revenue")
        ConsolidateColumn(cROAS, sHashPre & "ROAS")
        ConsolidateColumn(cCost, sHashPre & "Cost")
        ConsolidateColumn(cAssists, sHashPre & "Assists")
        ConsolidateColumn(cViewThroughConv, sHashPre & "ViewThroughConv")
        ConsolidateColumn(cConvOneClick, sHashPre & "ConvOneClick")
        ConsolidateColumn(cConvValue, sHashPre & "ConvValue")
        ConsolidateColumn(cTrans, sHashPre & "Trans")
        AverageColumnV(cAvgPos, sHashPre & "AvgPos")

    End Sub
    '========================================================================================================
    'ConsolidateGA() - ConsolidateGA - calling ConsolidateColumn() and AverageColumn() for numeric data   
    ' Only called from UpdateSQLGA
    '========================================================================================================
    Private Sub ConsolidateGA()
        Dim ncount As Integer

        'Columns numbers for report
        Dim cCampaign As Integer = -1
        Dim cAdGroup As Integer = -1
        Dim cKeyword As Integer = -1
        Dim cImpressions As Integer = -1
        Dim cClicks As Integer = -1
        Dim cCTR As Integer = -1
        Dim cCPC As Integer = -1
        Dim cConversions As Integer = -1
        Dim cConvValue As Integer = -1
        Dim cViewThroughConv As Integer = -1
        Dim cConvOneClick As Integer = -1
        Dim cClickConvRate As Integer = -1
        Dim cCostPerConv As Integer = -1
        Dim cRevenue As Integer = -1
        Dim cROAS As Integer = -1
        Dim cCost As Integer = -1
        Dim cAssists As Integer = -1
        Dim cAvgPos As Integer = -1

        Dim cTrans As Integer = -1

        'Locate columns assigning integer to variables
        For ncount = 0 To Columns.GetUpperBound(0)
            Select Case Columns(ncount).ToLower
                Case "campaign" : cCampaign = ncount
                Case "adgroup" : cAdGroup = ncount
                Case "keyword" : cKeyword = ncount
                Case "impressions" : cImpressions = ncount
                Case "clicks" : cClicks = ncount
                Case "ctr" : cCTR = ncount
                Case "cpc" : cCPC = ncount
                Case "conversions" : cConversions = ncount
                Case "viewthroughconv" : cViewThroughConv = ncount
                Case "convoneclick" : cConvOneClick = ncount
                Case "convvalue" : cConvValue = ncount
                Case "clickconvrate" : cClickConvRate = ncount
                Case "costperconv" : cCostPerConv = ncount
                Case "revenue" : cRevenue = ncount
                Case "roas" : cROAS = ncount
                Case "cost" : cCost = ncount
                Case "assists" : cAssists = ncount
                Case "avgpos" : cAvgPos = ncount
                Case "trans" : cTrans = ncount
            End Select
        Next

        'No data get out
        If cCampaign < 0 Then Exit Sub
        If cAdGroup < 0 Then Exit Sub
        If cKeyword < 0 Then Exit Sub

        'Verbiage for string columns assigned
        Dim sCampaign As String = Data(cCampaign)
        Dim sAdGroup As String = Data(cAdGroup)
        Dim sKeyword As String = Data(cKeyword)
        Dim sHashPre As String
        'GA data for adCenter has (Not Set) for AdGroup value, so eliminating when consolidating
        ' GetWhereGA()->GetDataForGA("AdGroup") will return a ""  
        If Me.EngineID = 3 And sAdGroup.Contains("(not set)") Then
            sHashPre = myDate.ToString("yyyyMMdd") & "." & sCampaign & "." & sKeyword & "."
        Else
            sHashPre = myDate.ToString("yyyyMMdd") & "." & sCampaign & "." & sAdGroup & "." & sKeyword & "."
        End If


        'Perform column consolidation for numeric columns
        ConsolidateColumn(cImpressions, sHashPre & "Impressions")
        ConsolidateColumn(cClicks, sHashPre & "Clicks")
        ConsolidateColumn(cCTR, sHashPre & "CTR")
        ConsolidateColumn(cCPC, sHashPre & "CPC")
        ConsolidateColumn(cConversions, sHashPre & "Conversions")
        ConsolidateColumn(cClickConvRate, sHashPre & "ClickConversionRate")
        ConsolidateColumn(cCostPerConv, sHashPre & "CostPerConv")
        ConsolidateColumn(cRevenue, sHashPre & "Revenue")
        ConsolidateColumn(cROAS, sHashPre & "ROAS")
        ConsolidateColumn(cCost, sHashPre & "Cost")
        ConsolidateColumn(cAssists, sHashPre & "Assists")
        ConsolidateColumn(cViewThroughConv, sHashPre & "ViewThroughConv")
        ConsolidateColumn(cConvOneClick, sHashPre & "ConvOneClick")
        ConsolidateColumn(cConvValue, sHashPre & "ConvValue")
        ConsolidateColumn(cTrans, sHashPre & "Trans")
        AverageColumnV(cAvgPos, sHashPre & "AvgPos")

    End Sub
    '========================================================================================================
    'ConsolidateColumn() - Consolidate data for numeric columns                                              
    '========================================================================================================
    Private Sub ConsolidateColumn(ByVal nID As Integer, ByVal sHashName As String)
        If nID < 0 Then Exit Sub

        If IsNumeric(Data(nID)) Then
            If IsNothing(myHash(sHashName)) Then
                myHash(sHashName) = Convert.ToDouble(Data(nID))
            Else
                Dim nTemp As Double = myHash(sHashName)
                nTemp = nTemp + Convert.ToDouble(Data(nID))
                myHash(sHashName) = nTemp
                Data(nID) = nTemp.ToString
            End If
        Else
            If Not IsNothing(myHash(sHashName)) Then
                Data(nID) = myHash(sHashName).ToString
            End If
        End If
    End Sub
    '========================================================================================================
    'AverageColumn() - For numeric data Average                                                              
    '========================================================================================================
    Private Sub AverageColumn(ByVal nid As Integer, ByVal sHashName As String)
        If nid < 0 Then Exit Sub

        If IsNumeric(Data(nid)) Then
            If IsNothing(myHash(sHashName)) Then
                myHash(sHashName) = Convert.ToDouble(Data(nid))
            Else
                Dim nTemp1 As Double = myHash(sHashName)
                Dim nTemp2 As Double = Convert.ToDouble(Data(nid))
                If nTemp1 > nTemp2 Then
                    myHash(sHashName) = nTemp1
                Else
                    myHash(sHashName) = nTemp2
                End If
                Data(nid) = myHash(sHashName).ToString
            End If
        Else
            If Not IsNothing(myHash(sHashName)) Then
                Data(nid) = myHash(sHashName).ToString
            End If
        End If
    End Sub
    '========================================================================================================
    'AverageColumn() - For numeric data Average                                                              
    '========================================================================================================
    Private Sub AverageColumnV(ByVal nid As Integer, ByVal sHashName As String)
        If nid < 0 Then Exit Sub

        If IsNumeric(Data(nid)) Then
            If IsNothing(myHash(sHashName)) Then
                myHash(sHashName) = Convert.ToDouble(Data(nid))
            Else
                Dim nTemp1 As Double = (myHash(sHashName) + Convert.ToDouble(Data(nid))) / 2
                myHash(sHashName) = nTemp1
            End If
            Data(nid) = myHash(sHashName).ToString

        Else
            If Not IsNothing(myHash(sHashName)) Then
                Data(nid) = myHash(sHashName).ToString
            End If
        End If
    End Sub
    '========================================================================================================
    'myRowGA() - Determine if data exist for the Campaign/AdGroup/Keyword for the Client/Engine/Date for GA  
    '========================================================================================================
    Private Function myRowGA(ByVal isArbitrage As Integer) As Integer
        Dim SQL As String
        If isArbitrage = 1 Then
            SQL = "SELECT RecID " & _
                  "FROM PerfDataArbitrage " & _
                    GetWhereGA()
        Else
            SQL = "SELECT RecID " & _
                  "FROM PerfData " & _
                    GetWhereGA()
        End If

        Dim conn As New SqlConnection(cs)
        conn.Open()

        Dim cmd As New SqlCommand(SQL, conn)
        Dim nAccount As Integer = cmd.ExecuteScalar

        cmd.Dispose()
        conn.Close()
        conn.Dispose()

        Return nAccount
    End Function
    'GA Visitor Data
    '========================================================================================================
    'ConsolidateV() - ConsolidateV - calling ConsolidateColumn() and AverageColumn() for numeric data          
    '========================================================================================================
    Private Sub ConsolidateV()
        Dim ncount As Integer

        'Columns numbers for report
        Dim cCustomVarName As Integer = -1
        Dim cCustomVarValue As Integer = -1

        Dim cVisits As Integer = -1
        Dim cPageViews As Integer = -1
        Dim cPageViewsPerVisit As Integer = -1
        Dim cAvgTimeOnSite As Integer = -1
        Dim cPerCentNewVisits As Integer = -1
        Dim cVisitBounceRate As Integer = -1
        Dim cGoal3Completions As Integer = -1
        Dim cGoal16Completions As Integer = -1
  

        'Locate columns assigning integer to variables
        For ncount = 0 To Columns.GetUpperBound(0)
            Select Case Columns(ncount).ToLower
                Case "customvarname" : cCustomVarName = ncount
                Case "customvarvalue" : cCustomVarValue = ncount
                Case "visits" : cVisits = ncount
                Case "pageviews" : cPageViews = ncount
                Case "pageviewpervisist" : cPageViewsPerVisit = ncount
                Case "avgtimeonsite" : cAvgTimeOnSite = ncount
                Case "percentnewvisits" : cPerCentNewVisits = ncount
                Case "visitbouncerate" : cVisitBounceRate = ncount
                Case "goal3completions" : cGoal3Completions = ncount
                Case "goal16completions" : cGoal16Completions = ncount
            End Select
        Next

        'No data get out
        If cCustomVarName < 0 Then Exit Sub
     

        'Verbiage for string columns assigned
        Dim sCustomVarName As String = Data(cCustomVarName)
        Dim sCustomVarValue As String = Data(cCustomVarValue)

        Dim sHashPre As String = myDate.ToString("yyyyMMdd") & "." & sCustomVarName & "." & sCustomVarValue & "."

        'Perform column consolidation for numeric columns
        ConsolidateColumn(cVisits, sHashPre & "Visits")
        ConsolidateColumn(cPageViews, sHashPre & "PageViews")
        ConsolidateColumn(cPageViewsPerVisit, sHashPre & "PageViewsPerVisit")
        ConsolidateColumn(cAvgTimeOnSite, sHashPre & "AvgTimeOnSite")
        ConsolidateColumn(cGoal3Completions, sHashPre & "ScheduleAppt")
        ConsolidateColumn(cGoal16Completions, sHashPre & "PrintQuote")
        AverageColumnV(cPerCentNewVisits, sHashPre & "PerCentNewVisits")
        AverageColumnV(cVisitBounceRate, sHashPre & "VisitBounceRate")
    End Sub


    '========================================================================================================
    'myRowGAV() - Determine if data exist for the Campaign/AdGroup/Keyword for the Client/Engine/Date for GA  
    '========================================================================================================
    Private Function myRowGAV(ByVal isArbitrage As Integer) As Integer
        Dim SQL As String
        If isArbitrage = 1 Then
            SQL = "SELECT RecID " & _
                  "FROM TireChoiceGAData " & _
                    GetWhereGAV()
        Else
            SQL = "SELECT RecID " & _
                  "FROM TireChoiceGAData " & _
                    GetWhereGAV()
        End If

        Dim conn As New SqlConnection(cs)
        conn.Open()

        Dim cmd As New SqlCommand(SQL, conn)
        Dim nAccount As Integer = cmd.ExecuteScalar

        cmd.Dispose()
        conn.Close()
        conn.Dispose()

        Return nAccount
    End Function
    '========================================================================================================
    'myRowFlexEOM() - Determine if data exist for the Campaign/AdGroup/Keyword for the Client/Engine/Date    
    '========================================================================================================
    Private Function myRowFlexEOM(ByVal isArbitrage As Integer, ByVal cmd As SqlCommand) As Integer
        Dim SQL As String
        Dim nAccount As Integer = 0
        Dim nConv As Integer = 0
        Dim nConvPrev As Integer = 0
        Dim nCount As Integer = 0
        If isArbitrage = 1 Then
            SQL = "SELECT RecID, Conversions " & _
                  "FROM PerfDataArbitrage " & _
                   GetWhere()
        Else
            SQL = "SELECT RecID, Conversions " & _
                  "FROM PerfData " & _
                   GetWhere()
        End If
        cmd.CommandText = SQL

        ''April 26 2011 for efficiency
        ''Dim conn As New SqlConnection(cs)
        ''conn.Open()
        ''Dim cmd As New SqlCommand(SQL, conn)
        ''cmd.CommandTimeout = 0
        Dim rdr As SqlDataReader = cmd.ExecuteReader

        nAccount = 0
        nConv = 0
        nConvPrev = 0
        If rdr.Read Then
            nAccount = rdr("RecID")
            If Not IsDBNull(rdr("Conversions")) Then nConvPrev = Convert.ToInt32(rdr("Conversions"))
            'Now search current reports data and determine if conversion value changed
            If Not IsNothing(Columns) Then
                For nCount = Columns.GetLowerBound(0) To Columns.GetUpperBound(0)
                    If Not IsNothing(Columns(nCount)) Then
                        If Columns(nCount).ToUpper = "CONVERSIONS" Then
                            nConv = Data(nCount)
                            Exit For
                        End If
                    End If
                Next

            End If
            If nConv <= nConvPrev Then nAccount = 0
        End If
        ''cmd.Dispose()
        ''conn.Close()
        ''conn.Dispose()
        rdr.Close()

        Return nAccount
    End Function
    '========================================================================================================
    'myRow() - Determine if data exist for the Campaign/AdGroup/Keyword for the Client/Engine/Date           
    '========================================================================================================
    Private Function myRow(ByVal isArbitrage As Integer, ByVal cmd As SqlCommand) As Integer
        Dim SQL As String
        If isArbitrage = 1 Then
            SQL = "SELECT RecID " & _
                  "FROM PerfDataArbitrage " & _
                   GetWhere()
        Else
            SQL = "SELECT RecID " & _
                  "FROM PerfData " & _
                   GetWhere()
        End If
        cmd.CommandText = SQL

        ''April 26 2011 for efficiency
        ''Dim conn As New SqlConnection(cs)
        ''conn.Open()
        ''Dim cmd As New SqlCommand(SQL, conn)
        ''cmd.CommandTimeout = 0

        Dim nAccount As Integer = cmd.ExecuteScalar

        ''cmd.Dispose()
        ''conn.Close()
        ''conn.Dispose()

        Return nAccount
    End Function
    '========================================================================================================
    'GetColumns() - Build columns                                                                            
    '========================================================================================================
    Private Function GetColumns() As String
        Dim sResult As String = "[Date],CustomerID,EngineID"
        Dim nCount As Integer
        For nCount = Columns.GetLowerBound(0) To Columns.GetUpperBound(0)
            If Not IsNothing(Columns(nCount)) Then
                If (Columns(nCount).ToLower <> "engine") Then
                    If Columns(nCount).Length > 0 Then sResult = sResult & "," & Columns(nCount)
                End If
            End If
        Next
        sResult = "(" & sResult & ")"

        Return sResult
    End Function
#End Region 'UpdateSQL Function and Such

#Region " AddRow and Such "
    '========================================================================================================
    'AddRowRegular() - Perfdata Add Data                                                                            
    '========================================================================================================
    Private Function AddRowRegular(ByVal cmd As SqlCommand) As Boolean
        Dim bResult As Boolean = True

        Dim strColumns As String = GetColumns()
        Dim strValues As String = GetAddData()

        Dim SQL As String
        SQL = "INSERT INTO PerfData " & strColumns & " " & _
              "VALUES " & strValues

        Try
            cmd.CommandText = SQL
            Dim nResult As Integer = cmd.ExecuteScalar

        Catch ex As Exception
            MsgBox("AddRow Error: " & ex.Message)
            bResult = False
        End Try

        Return bResult
    End Function
    '========================================================================================================
    'AddRowRegular() - Perfdata Add Data                                                                            
    '========================================================================================================
    Private Function AddRowRegularV(ByVal cmd As SqlCommand) As Boolean
        Dim bResult As Boolean = True

        Dim strColumns As String = GetColumns()
        Dim strValues As String = GetAddData()

        Dim SQL As String
        SQL = "INSERT INTO TireChoiceGAData " & strColumns & " " & _
              "VALUES " & strValues

        Try
            cmd.CommandText = SQL
            Dim nResult As Integer = cmd.ExecuteScalar

        Catch ex As Exception
            MsgBox("AddRow Error: " & ex.Message)
            bResult = False
        End Try

        Return bResult
    End Function
    '========================================================================================================
    'AddRow() - Perfdata Add Data                                                                            
    '========================================================================================================
    Private Function AddRowArbitrage(ByVal cmd As SqlCommand) As Boolean
        Dim bResult As Boolean = True

        Dim strColumns As String = GetColumns()
        Dim strValues As String = GetAddData()

        Dim SQL As String
        SQL = "INSERT INTO PerfDataArbitrage " & strColumns & " " & _
              "VALUES " & strValues

        Try
            cmd.CommandText = SQL
            Dim nResult As Integer = cmd.ExecuteScalar
        Catch ex As Exception
            MsgBox("AddRow Error: " & ex.Message)
            bResult = False
        End Try

        Return bResult
    End Function

    '========================================================================================================
    'AddRow() - Perfdata Add Data                                                                            
    '========================================================================================================
    Private Function AddRowArbitrageV(ByVal cmd As SqlCommand) As Boolean
        Dim bResult As Boolean = True

        Dim strColumns As String = GetColumns()
        Dim strValues As String = GetAddData()

        Dim SQL As String
        SQL = "INSERT INTO TireChoiceGAData " & strColumns & " " & _
              "VALUES " & strValues

        Try
            cmd.CommandText = SQL
            Dim nResult As Integer = cmd.ExecuteScalar
        Catch ex As Exception
            MsgBox("AddRow Error: " & ex.Message)
            bResult = False
        End Try

        Return bResult
    End Function

    '========================================================================================================
    'GetAddData() - Build record for Perfdata Add                                                            
    '========================================================================================================
    Private Function GetAddData() As String
        Dim sResult As String = "'" & myDate.ToString("yyyyMMdd") & "'," & CustomerID.ToString & "," & EngineID.ToString
        Dim nCount As Integer
        Dim sTemp As String
        Dim sType As String

        For nCount = Columns.GetLowerBound(0) To Columns.GetUpperBound(0)
            If Not IsNothing(Columns(nCount)) Then
                If Columns(nCount).Length > 0 Then
                    If Columns(nCount).ToLower <> "engine" Then
                        sType = DataTypes(Columns(nCount).ToLower)

                        Select Case sType.ToUpper
                            Case "STRING"
                                sTemp = Replace(Data(nCount), "'", "''")
                                sResult = sResult & ",'" & sTemp & "'"
                            Case Else
                                If Data(nCount).Length > 0 Then
                                    If IsNumeric(Data(nCount)) Then
                                        sResult = sResult & "," & Data(nCount)
                                    Else
                                        sResult = sResult & ",NULL"
                                    End If
                                Else
                                    sResult = sResult & ",NULL"
                                End If
                        End Select
                    End If
                End If
            End If
        Next
        sResult = "(" & sResult & ")"

        Return sResult
    End Function
#End Region 'AddRow and Such

#Region " UpdateRow and Such "
    '========================================================================================================
    'UpdateRowArbitrage(nRow) - Perfdata Update for record ID nRow passed                                       
    '========================================================================================================
    Private Function UpdateRowArbitrage(ByVal nRow As Integer, ByVal cmd As SqlCommand) As Boolean
        Dim bResult As Boolean = True
        Dim SQL As String

        SQL = "UPDATE PerfdataArbitrage " & _
              "SET " & GetUpdateData() & ",modifiedcost=NULL " & _
              "WHERE RecID = " & nRow.ToString

        Try
            cmd.CommandText = SQL
            Dim nResult As Integer = cmd.ExecuteScalar

        Catch ex As Exception
            Debug.Print("UPdate Error: " & ex.Message)
            bResult = False
        End Try

        Return bResult
    End Function
    '========================================================================================================
    'UpdateRowArbitrage(nRow) - Perfdata Update for record ID nRow passed                                       
    '========================================================================================================
    Private Function UpdateRowArbitrageV(ByVal nRow As Integer, ByVal cmd As SqlCommand) As Boolean
        Dim bResult As Boolean = True
        Dim SQL As String

        SQL = "UPDATE TireChoiceGAData " & _
              "SET " & GetUpdateDataV() & ",modifiedcost=NULL " & _
              "WHERE RecID = " & nRow.ToString

        Try
            cmd.CommandText = SQL
            Dim nResult As Integer = cmd.ExecuteScalar

        Catch ex As Exception
            Debug.Print("UPdate Error: " & ex.Message)
            bResult = False
        End Try

        Return bResult
    End Function
    '========================================================================================================
    'UpdateRowArbitrage(nRow) - Perfdata Update for record ID nRow passed                                       
    '========================================================================================================
    Private Function UpdateRowArbitrageFlexEOM(ByVal nRow As Integer, ByVal cmd As SqlCommand) As Boolean
        Dim bResult As Boolean = True
        Dim SQL As String

        SQL = "UPDATE PerfdataArbitrage " & _
              "SET " & GetUpdateDataFlexEOM() & ",modifiedcost=NULL " & _
              "WHERE RecID = " & nRow.ToString

        Try
            cmd.CommandText = SQL
            Dim nResult As Integer = cmd.ExecuteScalar

        Catch ex As Exception
            Debug.Print("UPdate Error: " & ex.Message)
            bResult = False
        End Try

        Return bResult
    End Function

    '========================================================================================================
    'UpdateRowRegular(nRow) - Perfdata Update for record ID nRow passed                                       
    '========================================================================================================
    Private Function UpdateRowRegular(ByVal nRow As Integer, ByVal cmd As SqlCommand) As Boolean
        Dim bResult As Boolean = True
        Dim SQL As String

        SQL = "UPDATE Perfdata " & _
              "SET " & GetUpdateData() & " " & _
              "WHERE RecID = " & nRow.ToString

        Try
            cmd.CommandText = SQL
            Dim nResult As Integer = cmd.ExecuteScalar
        Catch ex As Exception
            Debug.Print("UPdate Error: " & ex.Message)
            bResult = False
        End Try

        Return bResult
    End Function

    '========================================================================================================
    'UpdateRowRegular(nRow) - Perfdata Update for record ID nRow passed                                       
    '========================================================================================================
    Private Function UpdateRowRegularV(ByVal nRow As Integer, ByVal cmd As SqlCommand) As Boolean
        Dim bResult As Boolean = True
        Dim SQL As String

        SQL = "UPDATE TireChoiceGAData " & _
              "SET " & GetUpdateDataV() & " " & _
              "WHERE RecID = " & nRow.ToString

        Try
            cmd.CommandText = SQL
            Dim nResult As Integer = cmd.ExecuteScalar
        Catch ex As Exception
            Debug.Print("UPdate Error: " & ex.Message)
            bResult = False
        End Try

        Return bResult
    End Function
    '========================================================================================================
    'GetUpdateData() - Build record for Perfdata Update                                                      
    '========================================================================================================
    Private Function GetUpdateData() As String
        Dim sResult As String = ""
        Dim nCount As Integer
        Dim sTemp As String
        Dim sType As String

        For nCount = Columns.GetLowerBound(0) To Columns.GetUpperBound(0)
            If Not IsNothing(Columns(nCount)) Then
                If Columns(nCount).Length > 0 Then
                    If Columns(nCount).ToUpper <> "KEYWORD" Then
                        If Columns(nCount).ToUpper <> "ADGROUP" Then
                            If Columns(nCount).ToUpper <> "CAMPAIGN" Then
                                If Columns(nCount).ToUpper <> "ENGINE" Then
                                    sType = DataTypes(Columns(nCount).ToLower)
                                    Select Case sType.ToUpper
                                        Case "STRING"
                                            sTemp = Replace(Data(nCount), "'", "''")
                                            sResult = sResult & "," & Columns(nCount) & "='" & sTemp & "'"
                                        Case Else
                                            If Data(nCount).Length > 0 Then
                                                If IsNumeric(Data(nCount)) Then
                                                    sResult = sResult & "," & Columns(nCount) & "=" & Data(nCount).ToString
                                                Else
                                                    sResult = sResult & "," & Columns(nCount) & "=NULL"
                                                End If
                                            Else
                                                sResult = sResult & "," & Columns(nCount) & "=NULL"
                                            End If
                                    End Select
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        sResult = Mid(sResult, 2)

        Return sResult
    End Function
    '========================================================================================================
    'GetUpdateData() - Build record for Perfdata Update                                                      
    '========================================================================================================
    Private Function GetUpdateDataV() As String
        Dim sResult As String = ""
        Dim nCount As Integer
        Dim sTemp As String
        Dim sType As String

        For nCount = Columns.GetLowerBound(0) To Columns.GetUpperBound(0)
            If Columns(nCount).Length > 0 Then
                If Columns(nCount).ToUpper <> "CUSTOMVARNAME1" Then
                    If Columns(nCount).ToUpper <> "CUSTOMVARVALUE1" Then
                        If Columns(nCount).ToUpper <> "ENGINE" Then
                            sType = DataTypes(Columns(nCount).ToLower)
                            Select Case sType.ToUpper
                                Case "STRING"
                                    sTemp = Replace(Data(nCount), "'", "''")
                                    sResult = sResult & "," & Columns(nCount) & "='" & sTemp & "'"
                                Case Else
                                    If Data(nCount).Length > 0 Then
                                        If IsNumeric(Data(nCount)) Then
                                            sResult = sResult & "," & Columns(nCount) & "=" & Data(nCount).ToString
                                        Else
                                            sResult = sResult & "," & Columns(nCount) & "=NULL"
                                        End If
                                    Else
                                        sResult = sResult & "," & Columns(nCount) & "=NULL"
                                    End If
                            End Select
                        End If
                    End If
                End If
            End If
        Next
        sResult = Mid(sResult, 2)

        Return sResult
    End Function
    '========================================================================================================
    'GetUpdateDataFlexEOM() - Build record for Perfdata Update                                                      
    '========================================================================================================
    Private Function GetUpdateDataFlexEOM() As String
        Dim sResult As String = ""
        Dim nCount As Integer
        Dim sTemp As String
        Dim sType As String
     
        For nCount = Columns.GetLowerBound(0) To Columns.GetUpperBound(0)
            If Columns(nCount).Length > 0 Then
                If Columns(nCount).ToUpper <> "KEYWORD" Then
                    If Columns(nCount).ToUpper <> "ADGROUP" Then
                        If Columns(nCount).ToUpper <> "CAMPAIGN" Then
                            If Columns(nCount).ToUpper <> "ENGINE" Then
                                If Columns(nCount).ToUpper = "CONVERSIONS" Or Columns(nCount).ToUpper = "CLICKCONVRATE" Or Columns(nCount).ToUpper = "COSTPERCONV" Or Columns(nCount).ToUpper = "VIEWTHROUGHCONV" Then
                                    sType = DataTypes(Columns(nCount).ToLower)
                                    Select Case sType.ToUpper
                                        Case "STRING"
                                            sTemp = Replace(Data(nCount), "'", "''")
                                            sResult = sResult & "," & Columns(nCount) & "='" & sTemp & "'"
                                        Case Else
                                            If Data(nCount).Length > 0 Then
                                                If IsNumeric(Data(nCount)) Then
                                                    sResult = sResult & "," & Columns(nCount) & "=" & Data(nCount).ToString
                                                Else
                                                    sResult = sResult & "," & Columns(nCount) & "=NULL"
                                                End If
                                            Else
                                                sResult = sResult & "," & Columns(nCount) & "=NULL"
                                            End If
                                    End Select
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        sResult = Mid(sResult, 2)

        Return sResult
    End Function
#End Region 'UpdateRow and Such

#Region " Public Sub GetClientList "
    '===================================================================================================================================
    'GetClientList() Returning Clients() for frmMain to display and to pass along to all the methods used to communicate with the APIs  
    'Build query over Customers and X_Client_Campaign to retrieve active clients and engines information                                
    '===================================================================================================================================
    Public Function GetClientList() As Clients()

        Dim sGGLClient_ID As String = "(SELECT GGLClientID FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=2) AND (Status=1))) AS GGL_ID"

        Dim sYSMV6ID_ID As String = "(SELECT YSMV6ID FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=1) AND (Status=1))) AS YSMV6_ID"

        Dim sYahoo_ID As String = "(SELECT AccountID FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=1) AND (Status=1))) AS Yahoo_ID"
        Dim sYahoo_Sub As String = "(SELECT MemberOfMaster FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=1) AND (Status=1))) AS Yahoo_Sub"
        Dim sYahoo_Pwd As String = "(SELECT AccountPassword FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=1) AND (Status=1))) AS Yahoo_PWD"
        Dim sYahoo_GroupAcct As String = "(SELECT MasterAccountID FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=1) AND (Status=1))) AS Yahoo_GroupAcct"
        Dim sYahoo_GroupPwd As String = "(SELECT MasterAccountPassword FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=1) AND (Status=1))) AS Yahoo_GroupPWD"

        Dim sGoogle_ID As String = "(SELECT AccountID FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=2) AND (Status=1))) AS Google_ID"
        Dim sGoogle_Pwd As String = "(SELECT AccountPassword FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=2) AND (Status=1))) AS Google_PWD"
        Dim sGoogle_ReportID As String = "(SELECT ReportID FROM GoogleReportID WHERE (CustomerID=cust.CustomerID)) AS Google_ReportID"
        Dim sGoogle_EOMReportID As String = "(SELECT EOMReportID FROM GoogleReportID WHERE (CustomerID=cust.CustomerID)) AS Google_EOMReportID"

        Dim sMSN_ID As String = "(SELECT AccountID FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=3) AND (Status=1))) AS MSN_ID"
        Dim sMSN_Pwd As String = "(SELECT AccountPassword FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=3) AND (Status=1))) AS MSN_PWD"

        Dim sVerizon_ID As String = "(SELECT AccountID FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=4) AND (Status=1))) AS Verizon_ID"
        Dim sVerizon_Pwd As String = "(SELECT AccountPassword FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=4) AND (Status=1))) AS Verizon_PWD"

        Dim sASK_ID As String = "(SELECT AccountID FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=5) AND (Status=1))) AS ASK_ID"
        Dim sASK_Pwd As String = "(SELECT AccountPassword FROM X_Client_Campaign WHERE ((CustomerID=cust.CustomerID) AND (EngineID=5) AND (Status=1))) AS ASK_PWD"

        ''Dim SQL As String = "SELECT CustomerID, CustomerName, GAProfile, " & _
        ''                            sYahoo_ID & ", " & sYahoo_Pwd & ", " & sYahoo_Sub & ", " & _
        ''                            sYahoo_GroupAcct & ", " & sYahoo_GroupPwd & ", " & _
        ''                            sGoogle_ID & ", " & sGoogle_Pwd & ", " & _
        ''                            sMSN_ID & ", " & sMSN_Pwd & ", " & _
        ''                            sVerizon_ID & ", " & sVerizon_Pwd & ", " & _
        ''                            sASK_ID & ", " & sASK_Pwd & " " & _
        ''                    "FROM Customers cust " & _
        ''                    "WHERE Active=1 " & _
        ''                    "ORDER BY CustomerName"
        'Sherv Updated the line above to include Arbitrage customers

        ''Temp for V13 remove sGoogle_ReportID and sGoogle_EOMReportID work
        'Dim SQL As String = "SELECT CustomerID, CustomerName, GAProfile,isArbitrage, " & _
        '                            sYahoo_ID & ", " & sYahoo_Pwd & ", " & sYahoo_Sub & ", " & _
        '                            sYahoo_GroupAcct & ", " & sYahoo_GroupPwd & ", " & sYSMV6ID_ID & ", " & _
        '                            sGoogle_ID & ", " & sGoogle_Pwd & ", " & sGoogle_ReportID & ", " & sGoogle_EOMReportID & ", " & _
        '                            sMSN_ID & ", " & sMSN_Pwd & ", " & _
        '                            sVerizon_ID & ", " & sVerizon_Pwd & ", " & _
        '                            sASK_ID & ", " & sASK_Pwd & " " & _
        '                    "FROM Customers cust " & _
        '                    "WHERE (Active=1 or IsArbitrage=1)" & _
        '                    "ORDER BY CustomerName"

        Dim SQL As String = "SELECT CustomerID, CustomerName, GAProfile,isArbitrage, " & _
                            sYahoo_ID & ", " & sYahoo_Pwd & ", " & sYahoo_Sub & ", " & _
                            sYahoo_GroupAcct & ", " & sYahoo_GroupPwd & ", " & sYSMV6ID_ID & ", " & _
                            sGoogle_ID & ", " & sGoogle_Pwd & ", " & sGGLClient_ID & ", " & _
                            sMSN_ID & ", " & sMSN_Pwd & ", " & _
                            sVerizon_ID & ", " & sVerizon_Pwd & ", " & _
                            sASK_ID & ", " & sASK_Pwd & " " & _
                    "FROM Customers cust " & _
                    "WHERE (Active=1 or IsArbitrage=1)" & _
                    "ORDER BY CustomerName"

        Dim conn As New SqlConnection(cs)
        conn.Open()

        Dim cmd As New SqlCommand(SQL, conn)
        Dim rdr As SqlDataReader = cmd.ExecuteReader

        Dim MOM As Integer = 0
        Dim Index As Integer = 0
        Dim Clients() As Clients = Nothing

        'Loop through result set updating Clients variables
        While rdr.Read
            If IsNothing(Clients) Then
                Index = 0
                ReDim Clients(0)
            Else
                Index = Clients.GetUpperBound(0) + 1
            End If
            ReDim Preserve Clients(Index)
            Clients(Index) = New Clients

            With Clients(Index)
                'Sherv Added the following lines.
                If Not IsDBNull(rdr("IsArbitrage")) Then
                    .IsArbitrage = rdr("IsArbitrage")
                Else
                    .IsArbitrage = 0
                End If
                'Sherv Edits ended
                If Not IsDBNull(rdr("CustomerID")) Then
                    .CustID = rdr("CustomerID")
                End If

                If Not IsDBNull(rdr("CustomerName")) Then
                    .Name = rdr("CustomerName")
                End If

                If Not IsDBNull(rdr("GAProfile")) Then
                    .GA.AccountID = rdr("GAProfile")
                End If

                If Not IsDBNull(rdr("Google_ID")) Then
                    .Google.AccountID = rdr("Google_ID")
                End If

                If Not IsDBNull(rdr("Google_PWD")) Then
                    .Google.AccountPWD = rdr("Google_PWD")
                End If
                'Yahoo Group New April 1, 2010
                If Not IsDBNull(rdr("GGL_ID")) Then
                    .Google.GGLClientID = rdr("GGL_ID")
                End If
                ''Removed for V13
                'V201003 Added for Report Definition - Yesterday
                'If Not IsDBNull(rdr("Google_ReportID")) Then
                '    .Google.ReportID = rdr("Google_ReportID")
                'End If
                ''V201003 Added for Report Definition - EOM
                'If Not IsDBNull(rdr("Google_EOMReportID")) Then
                '    .Google.EOMReportID = rdr("Google_EOMReportID")
                'End If

                If Not IsDBNull(rdr("MSN_ID")) Then
                    .MSN.AccountID = rdr("MSN_ID")
                End If

                If Not IsDBNull(rdr("MSN_PWD")) Then
                    .MSN.AccountPWD = rdr("MSN_PWD")
                End If

                If Not IsDBNull(rdr("Verizon_ID")) Then
                    .Verizon.AccountID = rdr("Verizon_ID")
                End If

                If Not IsDBNull(rdr("Verizon_PWD")) Then
                    .Verizon.AccountPWD = rdr("Verizon_PWD")
                End If

                If Not IsDBNull(rdr("ASK_ID")) Then
                    .ASK.AccountID = rdr("ASK_ID")
                End If

                If Not IsDBNull(rdr("ASK_PWD")) Then
                    .ASK.AccountPWD = rdr("ASK_PWD")
                End If

                'Yahoo accounts under the umbrella MV account
                MOM = 0
                If Not IsDBNull(rdr("Yahoo_Sub")) Then
                    If IsNumeric(rdr("Yahoo_Sub")) Then
                        MOM = CType(rdr("Yahoo_Sub"), Integer)
                    End If
                End If

                If Not IsDBNull(rdr("Yahoo_ID")) Then
                    If MOM = 0 Then
                        .Yahoo.AccountID = rdr("Yahoo_ID")
                    Else
                        .Yahoo.AccountID = "214045"
                        .Yahoo.SubAccountID = rdr("Yahoo_ID")
                    End If
                End If

                If Not IsDBNull(rdr("Yahoo_PWD")) Then
                    .Yahoo.AccountPWD = rdr("Yahoo_PWD")
                End If
                'Yahoo Group New April 1, 2010
                If Not IsDBNull(rdr("YSMV6_ID")) Then
                    .Yahoo.YSMV6ID = rdr("YSMV6_ID")
                End If
                'Yahoo Group New March 23, 2009
                If Not IsDBNull(rdr("Yahoo_GroupAcct")) Then
                    If IsNumeric(rdr("Yahoo_GroupAcct")) Then
                    Else
                        .Yahoo.GroupAcct = rdr("Yahoo_GroupAcct")
                    End If
                End If
                If Not IsDBNull(rdr("Yahoo_GroupPWD")) Then
                    If IsNumeric(rdr("Yahoo_GroupPWD")) Then
                    Else
                        .Yahoo.GroupPWD = rdr("Yahoo_GroupPWD")
                    End If
                End If
            End With
        End While
        rdr.Close()
        cmd.Dispose()
        conn.Close()
        conn.Dispose()

        Return Clients
    End Function
    'V2010 Start
    '===================================================================================================================================
    'AddorUpdateGoogleReport() From frmMain to add/update GoogleReportID file with returned reportID for daily pulls of yesterday only                                
    '===================================================================================================================================
    Public Sub AddorUpdateGoogleReport(ByRef sCustomerID As String, ByRef sReportID As String, ByRef bEOM As Boolean)
        Dim SQL As String
        'Temporary for V13
        Return
        'Use myRow() to return row or zero designating update/add

        SQL = "SELECT CustomerID " & _
                          "FROM GoogleReportID WHERE CustomerID = " & sCustomerID

        Dim conn As New SqlConnection(cs)
        conn.Open()

        Dim cmd As New SqlCommand(SQL, conn)
        Dim nRow As Integer = cmd.ExecuteScalar

        cmd.Dispose()

        If nRow > 0 Then
            'Update Logic
            'Not EOM, so update ReportID instead of EOMReportID 
            If bEOM = False Then
                SQL = "UPDATE GoogleReportID " & _
                      "SET ReportID = " & _
                      "'" & sReportID & "'" & _
                      " WHERE CustomerID = " & nRow.ToString
            Else
                SQL = "UPDATE GoogleReportID " & _
                      "SET EOMReportID = " & _
                      "'" & sReportID & "'" & _
                      " WHERE CustomerID = " & nRow.ToString
            End If
        Else
            'Add Logic
            'Not EOM, so populate ReportID instead of EOMReportID 
            If bEOM = False Then
                SQL = "INSERT INTO GoogleReportID (CustomerID, ReportID) " & _
                       "VALUES (" & sCustomerID & ", '" & sReportID & "')"
            Else

                SQL = "INSERT INTO GoogleReportID (CustomerID, EOMReportID) " & _
                      "VALUES (" & sCustomerID & ", '" & sReportID & "')"
            End If
        End If

        Dim cmd2 As New SqlCommand(SQL, conn)
        cmd2.ExecuteScalar()
        cmd2.Dispose()
        conn.Close()
        conn.Dispose()
    End Sub
#End Region 'Public Sub GetClientList

#Region " Put in Engine ID Data Stuff NOT USED "
    '=========================================================================================================
    'UpdateX(Clients) - Determine if a client is in a particular engine and update or add to frmMain  
    'Not USED
    '=========================================================================================================
    Public Sub UpdateX(ByVal client As Clients)
        Dim sqlG As String = ""
        Dim sqlM As String = ""
        Dim sqlY As String = ""
        Dim sqlV As String = ""
        Dim sqlA As String = ""

        Dim Recid As Integer = -1

        'Determine if the client is in Yahoo
        If client.Yahoo.Has_AccountID Then
            Recid = RecID_X(1, client.CustID)
            If Recid > 0 Then
                'Update Yahoo
                Update_X(Recid, 1, client.Yahoo.AccountID, "")
            Else
                'Add Yahoo
                ADD_X(1, client.CustID, client.Yahoo.AccountID, "")
            End If
        End If

        'Determine if the client is in Google
        If client.Google.Has_AccountID Then
            Recid = RecID_X(2, client.CustID)
            If Recid > 0 Then
                'Update Google
                Update_X(Recid, 2, client.Google.AccountID, client.Google.AccountPWD)
            Else
                'Add Google
                ADD_X(2, client.CustID, client.Google.AccountID, client.Google.AccountPWD)
            End If
        End If

        'Determine if the client is in MSN
        If client.MSN.Has_AccountID Then
            Recid = RecID_X(3, client.CustID)
            If Recid > 0 Then
                'Update MSN
                Update_X(Recid, 3, client.MSN.AccountID, "")
            Else
                'Add MSN
                ADD_X(3, client.CustID, client.MSN.AccountID, "")
            End If
        End If

        'Determine if the client is in Verizon
        If client.Verizon.Has_AccountID Then
            Recid = RecID_X(4, client.CustID)
            If Recid > 0 Then
                'Update Verizon
                Update_X(Recid, 4, client.Verizon.AccountID, "")
            Else
                'Add Verizon
                ADD_X(4, client.CustID, client.Verizon.AccountID, "")
            End If
        End If

        'Determine if the client is in ASK
        If client.ASK.Has_AccountID Then
            Recid = RecID_X(5, client.CustID)
            If Recid > 0 Then
                'Update Verizon
                Update_X(Recid, 5, client.ASK.AccountID, "")
            Else
                'Add Verizon
                ADD_X(4, client.CustID, client.ASK.AccountID, "")
            End If
        End If
    End Sub

    Private Sub Update_X(ByVal RecID As Integer, ByVal EngineID As Integer, ByVal AccountID As String, ByVal Password As String)
        Dim SQL As String
        SQL = "UPDATE X_Client_Campaign " & _
              "SET AccountID = '" & AccountID & "', AccountPassword = '" & Password & "' " & _
              "WHERE TableID = " & RecID.ToString

        Try
            Dim conn As New SqlConnection(cs)
            conn.Open()

            Dim cmd As New SqlCommand(SQL, conn)
            Dim nResult As Integer = cmd.ExecuteScalar

            cmd.Dispose()
            conn.Close()
            conn.Dispose()

        Catch ex As Exception
            MsgBox("UPdate Error: " & ex.Message)
        End Try

    End Sub

    Private Sub ADD_X(ByVal EngineID As Integer, ByVal CustID As String, ByVal AccountID As String, ByVal Password As String)
        Dim sql As String

        sql = "INSERT INTO X_Client_Campaign (CustomerID,EngineID, Status, MemberOfMaster, AccountID, AccountPassword)" & _
              "VALUES (" & CustID & "," & EngineID.ToString & ",1," & "0" & ",'" & AccountID & "','" & Password & "')"

        Try
            Dim conn As New SqlConnection(cs)
            conn.Open()

            Dim cmd As New SqlCommand(sql, conn)
            Dim nResult As Integer = cmd.ExecuteScalar

            cmd.Dispose()
            conn.Close()
            conn.Dispose()

        Catch ex As Exception
            MsgBox("AddRow Error: " & ex.Message)
        End Try
    End Sub

    Private Function RecID_X(ByVal EngineID As Integer, ByVal CustID As String) As Integer
        Dim cmd As SqlCommand
        Dim nAcctID As Integer = 0

        Dim SQL As String
        SQL = "SELECT TableID " & _
              "FROM X_Client_Campaign " & _
              "WHERE CustomerID = " & CustID & " " & _
              "AND EngineID = " & EngineID.ToString

        Dim conn As New SqlConnection(cs)
        conn.Open()

        cmd = New SqlCommand(SQL, conn)
        nAcctID = cmd.ExecuteScalar

        cmd.Dispose()
        conn.Close()
        conn.Dispose()

        Return nAcctID
    End Function
#End Region 'Put in Engine ID Data Stuff

End Class

Public Class HashData
    Public Impressions As Double = 0
    Public Clicks As Double = 0
    Public CTR As Double = 0
    Public CPC As Double = 0
    Public Conversions As Double = 0
    Public ClickConvRate As Double = 0
    Public CostPerConv As Double = 0
    Public Revenue As Double = 0
    Public ROAS As Double = 0
    Public Cost As Double = 0
    Public Assists As Double = 0
    Public AvgPos As Double = 0
    Public ViewThroughConv As Double = 0
End Class