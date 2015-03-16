
Imports System.IO
Imports System.Data
Imports System.Xml
Imports System.Net
Imports System.Text
Imports System.Collections.Specialized
'JSON Parser Requirement
Imports System.Web.Script.Serialization


Namespace GAWebServices
    Public Class GAReports
        Public ws_AuthClient As AuthClient

        Private Const authOAFormat As String = "https://accounts.google.com/o/oauth2/token"
        Private Const authOARefreshToken As String = "refresh_token={0}&client_id={1}&client_secret={2}&grant_type=refresh_token"
        'Private Const authOABodyFormat As String = "code=4%2F5aMRZlCrzRGluNLeKb_H_hht_XbX&redirect_uri=urn:ietf:wg:oauth:2.0:oob&client_secret=V4KaSUgHS-F44-6M6V9mpcuh&grant_type=authorization_code&client_id=605424972061.apps.googleusercontent.com"

        Private Const authUrlFormat As String = "accountType=GOOGLE&Email={0}&Passwd={1}&source=MoreVisibility-GoogleAnalyticsReportInterface-1.0&service=analytics"
        'V3 Format
        Private Const requestUrlVFormat As String = "https://www.googleapis.com/analytics/v3/data/ga?ids={0}&dimensions={1}&metrics={2}&start-date={3}&end-date={4}&sort={5}&access_token={6}&start-index={7}&max-results=10000"
        Private Const requestUrlFormat As String = "https://www.googleapis.com/analytics/v3/data/ga?ids={0}&dimensions={1}&metrics={2}&filters=ga:medium%3D@cpc;ga:transactionRevenue%3E0&start-date={3}&end-date={4}&access_token={5}&start-index={6}&max-results=10000"

        'V2 Format
        'Private Const requestUrlVFormat As String = "https://www.google.com/analytics/feeds/data?ids={0}&dimensions={1}&metrics={2}&start-date={3}&end-date={4}&sort={5}"
        'Private Const requestUrlFormat As String = "https://www.google.com/analytics/feeds/data?ids={0}&dimensions={1}&metrics={2}&filters=ga:medium%3D@cpc;ga:transactionRevenue%3E0&start-date={3}&end-date={4}"
        ''Private Const requestUrlFormat As String = "https://www.google.com/analytics/feeds/data?ids={0}&dimensions={1}&metrics={2}&filters=ga:medium%3D%3Demail&start-date={3}&end-date={4}"

        Private Settings As New clSettings

        Public Email As String
        Public Password As String
        Public Token As String
        'V3 Added
        Public ClientID As String
        Public ClientSecret As String
        Public RefreshToken As String

        'Testing
        Public hshTable As New Hashtable()

        Private _BaseFilePath As String = ""
        Public DirectoryNamesAsRecID As Boolean = False

        Private idsDefault_Password As String
        Public Useragent As String

        Public AppToken As String

        Public AccountName As String

        Public StartDate As Date = Today
        Public EndDate As Date = Today

        Public Property BaseFilePath() As String
            Get
                Return _BaseFilePath
            End Get
            Set(ByVal value As String)
                If value.Length > 0 Then
                    If Right(value, 1) <> "\" Then value = value & "\"
                End If
                _BaseFilePath = value
            End Set
        End Property

        Private ReadOnly Property SavePath() As String
            Get
                Return BaseFilePath & "GA\"
            End Get
        End Property

#Region " Sub New "
        '========================================================================================
        'New(ByVal engineAuthentication As ClEngineSecurity)
        'Called by frmMain_Load()
        '========================================================================================
        Sub New(ByVal engineAuthentication As ClEngineSecurity)

            '
            'OAuth Refresh Token Work
            '
            ws_AuthClient = New AuthClient(engineAuthentication, True)
            Token = ws_AuthClient.GetAuthHeader()

        End Sub
        '========================================================================================
        'New(ByVal engineAuthentication As ClEngineSecurity)
        'Called by frmMain_Load()
        '========================================================================================
        Sub New(ByVal engineNames As clEngineNames)
            'Get GA authentication data from settings.xml 
            Settings = New clSettings
            Email = Settings.GAUser
            Password = Settings.GAPwd
            ClientID = Settings.GAClientID
            ClientSecret = Settings.GAClientSecret
            RefreshToken = Settings.GARefreshToken
            If Token = "" Then Token = GetToken(ClientID, ClientSecret, RefreshToken)

            hshTable = engineNames.HashEngines

            NewService()
        End Sub
        Public Function GetToken(ByVal ClientID As String, ByVal ClientSecret As String, ByVal RefreshToken As String) As String
            Dim strJSON As String
            Dim Tokens() As String
            Dim item As String
            Dim tbOutput As StringBuilder
            tbOutput = New StringBuilder()

            Dim indentLevel As Integer = 0

            'Dim authBody As String = String.Format(authUrlFormat, email, password)
            Dim authBody As String = String.Format(authOARefreshToken, RefreshToken, ClientID, ClientSecret)
            Dim req As HttpWebRequest
            Dim response As HttpWebResponse

            Dim ser As JavaScriptSerializer
            'Dim dict = New Dictionary(Of String, Object)()

            Dim stream As Stream
            Dim sw As StreamWriter
            Dim sr As StreamReader

            'V3 and OAuth2.0
            'End point for all Token work
            'Via Refresh Token, we get and Access Token and use for one hour.
            req = HttpWebRequest.Create("https://accounts.google.com/o/oauth2/token")
            req.Method = "POST"
            req.ContentType = "application/x-www-form-urlencoded"
            req.UserAgent = "GoogleAnalyticsReportInterface"

            stream = req.GetRequestStream()
            sw = New StreamWriter(stream)
            sw.Write(authBody)
            sw.Close()
            sw.Dispose()

            response = req.GetResponse()
            sr = New StreamReader(response.GetResponseStream())
            strJSON = sr.ReadToEnd()

            'JSON Work for finding access_token
            ser = New JavaScriptSerializer

            Dim dict As Dictionary(Of String, Object) = ser.Deserialize(Of Dictionary(Of String, Object))(strJSON)

            For Each strKey As String In dict.Keys
                If strKey.StartsWith("access_token") Then
                    Dim o As Object = dict(strKey)
                    Return o.ToString()
                End If
            Next
            'Dim objDeserialized As Object = ser.DeserializeObject(strJSON)


            'dict = ser.Deserialize(Of Dictionary(Of String, Object))(Token)
            'Tokens = Token.Split(ControlChars.CrLf.ToCharArray)
            'For Each item In Tokens
            '    If item.StartsWith("Auth=") Then
            '        Return item.Replace("Auth=", "")
            '    End If
            'Next

            Return String.Empty

        End Function

        Public Function GetToken(ByVal email As String, ByVal password As String) As String
            Dim Token As String
            Dim Tokens() As String
            Dim item As String

            Dim authBody As String = String.Format(authUrlFormat, email, password)
            Dim req As HttpWebRequest
            Dim response As HttpWebResponse

            Dim stream As Stream
            Dim sw As StreamWriter
            Dim sr As StreamReader

            req = HttpWebRequest.Create("https://www.google.com/accounts/ClientLogin")
            req.Method = "POST"
            req.ContentType = "application/x-www-form-urlencoded"
            req.UserAgent = "GoogleAnalyticsReportInterface"

            stream = req.GetRequestStream()
            sw = New StreamWriter(stream)
            sw.Write(authBody)
            sw.Close()
            sw.Dispose()

            response = req.GetResponse()
            sr = New StreamReader(response.GetResponseStream())
            Token = sr.ReadToEnd()
            Tokens = Token.Split(ControlChars.CrLf.ToCharArray)
            For Each item In Tokens
                If item.StartsWith("Auth=") Then
                    Return item.Replace("Auth=", "")
                End If
            Next

            Return String.Empty

        End Function
        Sub NewService()
            'to do
        End Sub
#End Region 'Sub New


#Region " My Functions invoked via frmMain Start_GA()"
        '=============================================================================================
        'ReqestReportID(ByVal client As Clients) called by frmMain.vb Start_GA()
        'Return String with ReportID or verbiage reflecting an error occurred while requesting: "ERROR"
        'Perform initialization 
        '=============================================================================================
        Public Function ReqestReports(ByVal client As Clients) As String
            'Check if token set to expire
            Token = ws_AuthClient.GetAuthHeader()

            Dim ReportID As String = "0"
            Dim myID As Integer = 0
            Dim myVID As Integer = 0
            'Multiple Profiles in comma seperated string profileID
            Dim profiles As StringBuilder = New StringBuilder()
            Dim sProfileIDs() As String = client.GA.AccountID.Split(",")
            Dim nCount As Integer
            myID = 0
            myVID = 0
            For nCount = sProfileIDs.GetLowerBound(0) To sProfileIDs.GetUpperBound(0)
                client.GA.AccountID = sProfileIDs(nCount).Trim(" ".ToCharArray())

                'Perform Report Build and Request for Revenue Report
                ReportID = ReqestReportIDRev(client)
                If ReportID = "GOOD" Then
                    myID = -1
                End If

                'Perform Report Build and Request for Custom Variable Report
                ReportID = ReqestReportID(client)
                If ReportID = "GOOD" Then
                    myVID = -1
                End If

            Next
            
            'Had problems with API - return an error indicator
            If myID = 0 And myVID = 0 Then Return "ERROR"

            Return "GOOD"

        End Function

        '=============================================================================================
        'ReqestReportIDRev(ByVal client As Clients) -->ReqestReports called by frmMain.vb Start_GA()
        'Return String with ReportID or verbiage reflecting an error occurred while requesting: "ERROR"
        'Perform initialization 
        '=============================================================================================
        Public Function ReqestReportIDRev(ByVal client As Clients) As String
            'JSON Parsing 
            Dim strJSON As String
            Dim ser As JavaScriptSerializer
            'Build File Name
            Dim sFile As String = BldFileName(client)
            'Start-index for result set
            Dim intTotRecs As Integer = 0
            Dim offset As Integer = 1

            Dim oRsp As String
            Dim myID As Long = -1
            Try
                oRsp = GetRevenueReport(Token, client.GA.AccountID, client.GA.StartDate, client.GA.EndDate, offset)

            Catch ex As Exception
                Log("GA ERROR ON " & client.LogName(True) & ex.Message)
                myID = 0

            End Try
            'Had problems with API - return an error indicator
            If myID = 0 Then Return "ERROR"

            'V3 New JSON Processing
            'If no errors begin parsing JSON results
            Dim sVar As String = String.Empty
            Dim sSource As String = String.Empty
            Dim sMedium As String = String.Empty
            Dim sCampaign As String = String.Empty
            Dim sAdGroup As String = String.Empty
            Dim sKeyWord As String = String.Empty
            Dim sLookup As String = String.Empty
            Dim dimension0 As String  'Date 
            Dim dimension1 As String  'Engine 
            Dim dimension2 As String  'Campaign
            Dim dimension3 As String  'AdGroup
            Dim dimension4 As String  'Keyword
            Dim metric As String      'Revenue
            Dim metricC As String      'Conversions
           
            'Writing out the raw JSON file
            Dim sw As StreamWriter
            sw = New StreamWriter(sFile)
            sw.Write(oRsp)
            sw.Dispose()

            'Write to file
            Dim objWriter As StreamWriter
            objWriter = New StreamWriter(sFile)
            Dim sLine As New StringBuilder()
            Dim bDeletedFile As Boolean = False

            'Build header line
            sLine.Append("Date,Engine,Campaign,AdGroup,Keyword,Revenue,Trans")
            sLine.AppendLine()
            objWriter.Write(sLine)
            sLine.Length = 0

            'Begin JSON Parsing of Result oRsp
            strJSON = oRsp
            ser = New JavaScriptSerializer
            Dim intItemsPerPage As Integer = 0
            Dim intTotalResults As Integer = 0


            Dim dict As Dictionary(Of String, Object) = ser.Deserialize(Of Dictionary(Of String, Object))(strJSON)
            'Start Loop for over number of records found in GA returned data sets based on returned totalResults and itemsPerPage
            Do
                'Find "rows" in JSON response which is now in Dictionary object dict
                For Each strKey As String In dict.Keys
                    'Get Items per Page and Total Results, to determine looping functions
                    If strKey.StartsWith("itemsPerPage") Then
                        intItemsPerPage = dict(strKey)
                    End If
                    If strKey.StartsWith("totalResults") Then
                        intTotalResults = dict(strKey)
                        'If we have no results, no output needded - move previously written raw file - sFile
                        'If intTotalResults = 0 Then
                        '    bDeletedFile = ChkMoveFile(client, sFile)
                        '    Exit Do
                        'End If
                    End If

                    'On with the Data
                    If strKey.StartsWith("rows") Then
                        Dim o As Object = dict(strKey)
                        If TypeOf o Is ArrayList Then
                            '=================================
                            'Process all Rows of Data in Array
                            For Each o2 As Object In DirectCast(o, ArrayList)
                                If TypeOf o2 Is ArrayList Then
                                    Dim intColumn As Integer = 0

                                    'Grab Row of data and move to sLine
                                    For Each oChild2 As Object In DirectCast(o2, ArrayList)
                                        If TypeOf oChild2 Is String Then
                                            Select Case intColumn
                                                Case 0
                                                    dimension0 = DirectCast(oChild2, String)

                                                Case 1
                                                    dimension1 = DirectCast(oChild2, String)

                                                Case 2
                                                    dimension2 = DirectCast(oChild2, String)

                                                Case 3
                                                    dimension3 = DirectCast(oChild2, String)

                                                Case 4
                                                    dimension4 = DirectCast(oChild2, String)

                                                Case 5
                                                    metric = DirectCast(oChild2, String)

                                                Case 6
                                                    metricC = DirectCast(oChild2, String)
                                            End Select
                                            intColumn += 1
                                        End If
                                    Next

                                    'Let's get the engineID from first dimension
                                    sVar = dimension1.Substring(0, 3).ToLower
                                    sLookup = dimension1.ToLower()

                                    'Skip YST
                                    If (dimension1.Substring(0, 3).ToLower <> "yst") Then

                                        Select Case sVar
                                            Case "yah"
                                                sVar = "3"
                                            Case "goo"
                                                sVar = "2"
                                            Case "msn"
                                                sVar = "3"
                                            Case "sup"
                                                sVar = "4"
                                            Case "ask"
                                                sVar = "5"
                                            Case "qui"
                                                sVar = "6"
                                            Case "bin"
                                                sVar = "3"
                                            Case "loo"
                                                sVar = "13"
                                            Case "fac"
                                                sVar = "17"
                                            Case "lin"
                                                sVar = "21"
                                            Case Else
                                                sVar = "2"
                                        End Select

                                        'Find Specific Source/Engines Names
                                        If hshTable.ContainsKey(sLookup) = True Then
                                            sVar = hshTable.Item(sLookup).ToString()
                                        End If
                                        'GA - no longer use Yahoo
                                        If sVar = "1" Then
                                            sVar = "3"
                                        End If

                                        'Date work
                                        sLine.Append(dimension0)
                                        sLine.Append(",")

                                        'Engine work
                                        sLine.Append(sVar)
                                        sLine.Append(",")

                                        'Campaign work
                                        sLine.Append(dimension2)
                                        sLine.Append(",")

                                        'AdGroup work
                                        sLine.Append(dimension3)
                                        sLine.Append(",")

                                        'Check if Keyword is for (content targeting)
                                        sVar = dimension4.ToLower
                                        If InStr(sVar, "(content targeting)") > 0 Then
                                            sVar = "Total - content targeting"
                                        End If
                                        sLine.Append(sVar)
                                        sLine.Append(",")

                                        'Revenue work $$ and # transactions
                                        sLine.Append(metric)
                                        sLine.Append(",")
                                        sLine.Append(metricC)

                                        sLine.AppendLine()
                                        objWriter.Write(sLine)
                                        sLine.Length = 0

                                    End If
                                End If
                            Next

                        End If
                    End If  ' Row Element
                Next        ' Process all JSON elements in Dictionary

                'Determine if we need to perform additional request from web service
                offset += intItemsPerPage
                If intTotalResults > offset Then
                    oRsp = GetRevenueReport(Token, client.GA.AccountID, client.GA.StartDate, client.GA.EndDate, offset)
                    'Begin JSON Parsing of Result oRsp
                    strJSON = oRsp
                    dict = ser.Deserialize(Of Dictionary(Of String, Object))(strJSON)
                End If
            Loop While (offset < intTotalResults)
                objWriter.Close()

                ''Testing via writing out the raw XML file
                'Dim sw As StreamWriter
                'sw = New StreamWriter(sFile)
                'sw.Write(oRsp)
                'sw.Dispose()

                'Dim resultsXml As XmlDocument
                'Dim entries As XmlNodeList
                'Dim i As Integer
                'Dim dimension1 As String  'Engine 
                'Dim dimension2 As String  'Campaign
                'Dim dimension3 As String  'AdGroup
                'Dim dimension4 As String  'Keyword
                'Dim metric As String      'Revenue

                'resultsXml = New XmlDocument
                'resultsXml.LoadXml(oRsp)
                'entries = resultsXml.GetElementsByTagName("entry")

                ''Write to file
                'Dim objWriter As StreamWriter
                'objWriter = New StreamWriter(sFile)
                'Dim sLine As New StringBuilder()


                ''Build header line
                'sLine.Append("Engine,Campaign,AdGroup,Keyword,Revenue")
                'sLine.AppendLine()
                'objWriter.Write(sLine)
                'sLine.Length = 0

                'Dim sVar As String = String.Empty
                'Dim sLookup As String

                ''Loop all Results returned from the API
                'For i = 0 To (entries.Count - 1)
                '    dimension1 = entries(i).ChildNodes(4).Attributes("value").Value
                '    dimension2 = entries(i).ChildNodes(5).Attributes("value").Value
                '    dimension3 = entries(i).ChildNodes(6).Attributes("value").Value
                '    dimension4 = entries(i).ChildNodes(7).Attributes("value").Value
                '    metric = entries(i).ChildNodes(8).Attributes("value").Value

                '    'Let's get the engineID from first dimension
                '    sVar = dimension1.Substring(0, 3).ToLower
                '    sLookup = dimension1.ToLower()

                '    'Skip YST
                '    If (dimension1.Substring(0, 3).ToLower <> "yst") Then

                '        Select Case sVar
                '            Case "yah"
                '                sVar = "3"
                '            Case "goo"
                '                sVar = "2"
                '            Case "msn"
                '                sVar = "3"
                '            Case "sup"
                '                sVar = "4"
                '            Case "ask"
                '                sVar = "5"
                '            Case "qui"
                '                sVar = "6"
                '            Case "bin"
                '                sVar = "3"
                '            Case "loo"
                '                sVar = "13"
                '            Case "fac"
                '                sVar = "17"
                '            Case "lin"
                '                sVar = "21"
                '            Case Else
                '                sVar = "2"
                '        End Select

                '        'Find Specific Source/Engines Names
                '        If hshTable.ContainsKey(sLookup) = True Then
                '            sVar = hshTable.Item(sLookup).ToString()
                '        End If
                '        'GA - no longer use Yahoo
                '        If sVar = "1" Then
                '            sVar = "3"
                '        End If

                '        'Build String
                '        sLine.Append(sVar)
                '        sLine.Append(",")

                '        'Campaign work
                '        sLine.Append(dimension2)
                '        sLine.Append(",")

                '        'AdGroup work
                '        sLine.Append(dimension3)
                '        sLine.Append(",")

                '        'Check if Keyword is for (content targeting)
                '        sVar = dimension4.ToLower
                '        If InStr(sVar, "(content targeting)") > 0 Then
                '            sVar = "Total - content targeting"
                '        End If
                '        sLine.Append(sVar)
                '        sLine.Append(",")

                '        'Revenue work
                '        sLine.Append(metric)

                '        sLine.AppendLine()
                '        objWriter.Write(sLine)
                '        sLine.Length = 0

                '    End If

                'Next

                'objWriter.Close()

                Return "GOOD"
        End Function
        '=============================================================================================
        'ReqestReportID(ByVal client As Clients) -->ReqestReports called by frmMain.vb Start_GA()
        'Return String with ReportID or verbiage reflecting an error occurred while requesting: "ERROR"
        'Perform initialization 
        '=============================================================================================
        Public Function ReqestReportID(ByVal client As Clients) As String
            'JSON Parsing 
            Dim strJSON As String
            Dim ser As JavaScriptSerializer
            'Build File Name
            Dim sFile As String = BldFileNameV(client)
            'Start-index for result set
            Dim intTotRecs As Integer = 0
            Dim offset As Integer = 1

            Dim oRsp As String
            Dim myID As Long = -1
            Try
                oRsp = GetVisitorReport(Token, client.GA.AccountID, client.GA.StartDate, client.GA.EndDate, offset)

            Catch ex As Exception
                Log("GA ERROR ON " & client.LogName(True) & ex.Message)
                myID = 0

            End Try
            'Had problems with API - return an error indicator
            If myID = 0 Then Return "ERROR"

            'Writing out the raw JSON file
            Dim sw As StreamWriter
            sw = New StreamWriter(sFile)
            sw.Write(oRsp)
            sw.Dispose()

            'Write to file
            Dim objWriter As StreamWriter
            objWriter = New StreamWriter(sFile)
            Dim sLine As New StringBuilder()

            'V3 New JSON Processing
            'If no errors begin parsing JSON results

            Dim i As Integer
            Dim dimension1 As String  'Custom Variable Name 1 
            Dim dimension2 As String  'Custom Variable Value 1
            Dim dimension3 As String  'Custom Variable Name 2 
            Dim dimension4 As String  'Custom Variable Value 2
            Dim metric1 As String      'visits
            Dim metric2 As String      'pageviews
            Dim metric3 As String      'pageviewPerVisit
            Dim metric4 As String      'avgTimeOnSite
            Dim metric5 As String      'percentNewVisits
            Dim metric6 As String      'visitBounceRate
            Dim metric7 As String      'goal3completions - Schedule Appt
            Dim metric8 As String      'goal16completions - Print Quote

            'Build header line
            'This must match exactly settings.xml GAVCoumn name attributes
            sLine.Append("CustomVarName1,CustomVarValue1,visits,pageviews,pageviewpervisist,avgtimeonsite,percentnewvisits,visitbouncerate,goal3completions,goal16completions")
            sLine.AppendLine()
            objWriter.Write(sLine)
            sLine.Length = 0

            Dim sVar As String = String.Empty
            Dim sLookup As String

            Dim intItemsPerPage As Integer = 0
            Dim intTotalResults As Integer = 0
            'Begin JSON Parsing of Result oRsp
            strJSON = oRsp
            ser = New JavaScriptSerializer

            Dim dict As Dictionary(Of String, Object) = ser.Deserialize(Of Dictionary(Of String, Object))(strJSON)
            intTotRecs = 0
            'Start Loop for over number of records found in GA returned data sets based on returned totalResults and itemsPerPage
            Do
                'Find "rows" in JSON response which is now in Dictionary object dict
                For Each strKey As String In dict.Keys
                    'Get Items per Page and Total Results, to determine looping functions
                    If strKey.StartsWith("itemsPerPage") Then
                        intItemsPerPage = dict(strKey)
                    End If
                    If strKey.StartsWith("totalResults") Then
                        intTotalResults = dict(strKey)
                        'If we have no results, no output needded - move previously written raw file - sFile
                        'If intTotalResults = 0 Then
                        '    bDeletedFile = ChkMoveFile(client, sFile)
                        '    Exit Do
                        'End If
                    End If

                    If strKey.StartsWith("rows") Then
                        Dim o As Object = dict(strKey)
                        If TypeOf o Is ArrayList Then
                            For Each o2 As Object In DirectCast(o, ArrayList)
                                If TypeOf o2 Is ArrayList Then
                                    Dim intColumn As Integer = 0
                                    For Each oChild2 As Object In DirectCast(o2, ArrayList)
                                        If TypeOf oChild2 Is String Then
                                            Select Case intColumn
                                                Case 0
                                                    dimension1 = DirectCast(oChild2, String)

                                                Case 1
                                                    dimension2 = DirectCast(oChild2, String)

                                                Case 2
                                                    metric1 = DirectCast(oChild2, String)

                                                Case 3
                                                    metric2 = DirectCast(oChild2, String)

                                                Case 4
                                                    metric3 = DirectCast(oChild2, String)

                                                Case 5
                                                    metric4 = DirectCast(oChild2, String)

                                                Case 6
                                                    metric5 = DirectCast(oChild2, String)

                                                Case 7
                                                    metric6 = DirectCast(oChild2, String)

                                                Case 8
                                                    metric7 = DirectCast(oChild2, String)

                                                Case 9
                                                    metric8 = DirectCast(oChild2, String)
                                            End Select
                                            intColumn += 1
                                        End If
                                    Next

                                    'Remove URL Encoding from string prior to using
                                    dimension1 = Replace(dimension1, ",", "-")
                                    dimension1 = System.Web.HttpUtility.UrlDecode(dimension1)
                                    dimension2 = Replace(dimension2, ",", "-")
                                    dimension2 = System.Web.HttpUtility.UrlDecode(dimension2)

                                    'Build String with Dimensions
                                    sVar = dimension1.Trim
                                    sLine.Append(sVar)
                                    sLine.Append(",")

                                    'Custom Variable Value 
                                    sVar = dimension2.Trim
                                    sLine.Append(sVar)
                                    sLine.Append(",")

                                    'Metrics work
                                    sLine.Append(metric1)
                                    sLine.Append(",")

                                    sLine.Append(metric2)
                                    sLine.Append(",")

                                    sLine.Append(metric3)
                                    sLine.Append(",")

                                    sLine.Append(metric4)
                                    sLine.Append(",")

                                    sLine.Append(metric5)
                                    sLine.Append(",")

                                    sLine.Append(metric6)
                                    sLine.Append(",")

                                    sLine.Append(metric7)
                                    sLine.Append(",")

                                    'Last one
                                    sLine.Append(metric8)

                                    sLine.AppendLine()
                                    objWriter.Write(sLine)
                                    sLine.Length = 0
                                End If
                            Next     'Row processing Loop

                        End If
                    End If
                Next                  'Process all JSON elements in Dictionary

                'Determine if we need to perform additional request from web service
                offset += intItemsPerPage
                If intTotalResults > offset Then
                    oRsp = GetRevenueReport(Token, client.GA.AccountID, client.GA.StartDate, client.GA.EndDate, offset)
                    'Begin JSON Parsing of Result oRsp
                    strJSON = oRsp
                    dict = ser.Deserialize(Of Dictionary(Of String, Object))(strJSON)
                End If
            Loop While (offset < intTotalResults)

            'V2 Logic for XML result set processing
            ''Loop all Results returned from the API
            'For i = 0 To (entries.Count - 1)
            '    dimension1 = entries(i).ChildNodes(4).Attributes("value").Value
            '    dimension2 = entries(i).ChildNodes(5).Attributes("value").Value
            '    metric1 = entries(i).ChildNodes(6).Attributes("value").Value
            '    metric2 = entries(i).ChildNodes(7).Attributes("value").Value
            '    metric3 = entries(i).ChildNodes(8).Attributes("value").Value
            '    metric4 = entries(i).ChildNodes(9).Attributes("value").Value
            '    metric5 = entries(i).ChildNodes(10).Attributes("value").Value
            '    metric6 = entries(i).ChildNodes(11).Attributes("value").Value
            '    metric7 = entries(i).ChildNodes(12).Attributes("value").Value
            '    metric8 = entries(i).ChildNodes(13).Attributes("value").Value

            '    'Remove URL Encoding from string prior to using
            '    dimension1 = Replace(dimension1, ",", "-")
            '    dimension1 = System.Web.HttpUtility.UrlDecode(dimension1)
            '    dimension2 = Replace(dimension2, ",", "-")
            '    dimension2 = System.Web.HttpUtility.UrlDecode(dimension2)

            '    'Build String with Dimensions
            '    sVar = dimension1.Trim
            '    sLine.Append(sVar)
            '    sLine.Append(",")

            '    'Custom Variable Value 
            '    sVar = dimension2.Trim
            '    sLine.Append(sVar)
            '    sLine.Append(",")

            '    'Metrics work
            '    sLine.Append(metric1)
            '    sLine.Append(",")

            '    sLine.Append(metric2)
            '    sLine.Append(",")

            '    sLine.Append(metric3)
            '    sLine.Append(",")

            '    sLine.Append(metric4)
            '    sLine.Append(",")

            '    sLine.Append(metric5)
            '    sLine.Append(",")

            '    sLine.Append(metric6)
            '    sLine.Append(",")

            '    sLine.Append(metric7)
            '    sLine.Append(",")

            '    'Last one
            '    sLine.Append(metric8)

            '    sLine.AppendLine()
            '    objWriter.Write(sLine)
            '    sLine.Length = 0

            'Next

            objWriter.Close()

            Return "GOOD"
        End Function
        '=============================================================================================
        'GetRevenueReport(ByVal sEMail As String, ByVal sPassword As String) called by ReqestReportID()
        'Instantiate and populate properties for myService via NewService():
        'myService = New GAWebServices.ReportService 
        'Used to invoke web services from GA for Revenue Metrics
        '=============================================================================================
        Public Function GetRevenueReport(ByVal Token As String, ByVal ProfileID As String, ByVal from As Date, ByVal todate As Date, ByVal offset As String) As String

            Dim authBody As String = String.Format(authUrlFormat, Email, Password)
            Dim myRequest As HttpWebRequest
            Dim myResponse As HttpWebResponse

            'Build Request URL in GetURI(Token, ProfileID, from, todate, offset)
            Dim uri As String
            uri = GetRevURI(Token, ProfileID, from, todate, offset)

            myRequest = HttpWebRequest.Create(uri)
            myRequest.Headers.Add("Authorization: GoogleLogin auth=" + Token)
            myResponse = myRequest.GetResponse

            Dim stream As Stream
            Dim sw As StreamWriter
            Dim sr As StreamReader
            Dim sResponse As String

            stream = myResponse.GetResponseStream()

            sr = New StreamReader(stream)
            sResponse = sr.ReadToEnd()

            Return sResponse
        End Function
        '=============================================================================================
        'GetVisitorReport(ByVal sEMail As String, ByVal sPassword As String) called by ReqestReportID()
        'Instantiate and populate properties for myService via NewService():
        'myService = New GAWebServices.ReportService 
        'Used to invoke web services from GA for Visitor Metrics
        '=============================================================================================
        Public Function GetVisitorReport(ByVal Token As String, ByVal ProfileID As String, ByVal from As Date, ByVal todate As Date, ByVal offset As String) As String

            Dim authBody As String = String.Format(authUrlFormat, Email, Password)
            Dim myRequest As HttpWebRequest
            Dim myResponse As HttpWebResponse

            'Build Request URL in GetURI(Token, ProfileID, from, todate, offset)
            Dim uri As String
            uri = GetVisitURI(Token, ProfileID, from, todate, offset)

            myRequest = HttpWebRequest.Create(uri)
            myRequest.Headers.Add("Authorization: GoogleLogin auth=" + Token)
            myResponse = myRequest.GetResponse

            Dim stream As Stream
            Dim sw As StreamWriter
            Dim sr As StreamReader
            Dim sResponse As String

            stream = myResponse.GetResponseStream()

            sr = New StreamReader(stream)
            sResponse = sr.ReadToEnd()

            Return sResponse
        End Function
        '=============================================================================================
        'GetRevURI--> Called by ReqestReports-->ReqestReportIDRev-->GetRevenueReport-->GetRevURI
        'Format URI with GA Profile; Dimensions(Source,Campaign,AdGroup,Keyword) & Metrics(Revenue,GoalCompletions)
        '=============================================================================================
        Public Function GetRevURI(ByVal sessionToken As String, ByVal profileID As String, ByVal from As Date, ByVal todate As Date, ByVal offset As String) As String
            Dim objItemD As Dimension
            Dim objItemM As Metric

            Dim objDimension() As Dimension = New Dimension() {Dimension.source, Dimension.campaign, Dimension.adGroup, Dimension.keyword}
            Dim objMetrics() As Metric = New Metric() {Metric.transactionRevenue, Metric.goal1Completions}

            Dim dims As StringBuilder = New StringBuilder()
            dims.Append("ga:date,")
            For Each objItemD In objDimension
                dims.Append("ga:" + objItemD.ToString() + ",")
            Next

            Dim mets As StringBuilder = New StringBuilder()
            For Each objItemM In objMetrics
                mets.Append("ga:" + objItemM.ToString() + ",")
            Next

            Dim requestUrl As String

            requestUrl = String.Format(requestUrlFormat, "ga:" + profileID, _
                                                         dims.ToString().Trim(",".ToCharArray()), _
                                                         mets.ToString().Trim(",".ToCharArray()), _
                                                         from.ToString("yyyy-MM-dd"), _
                                                         todate.ToString("yyyy-MM-dd"), _
                                                         sessionToken, offset)

            Return requestUrl
        End Function
        '=============================================================================================
        'GetVisitURI--> Called by ReqestReports-->ReqestReportID-->GetVisitorReport-->GetVisitURI
        'Format URI with GA Profile; Dimensions(CustomVariables) & Metrics(Visitor Data)
        '=============================================================================================
        Public Function GetVisitURI(ByVal sessionToken As String, ByVal profileID As String, ByVal from As Date, ByVal todate As Date, ByVal offset As String) As String
            Dim objItemD As Dimension
            Dim objItemM As Metric

            Dim objDimension() As Dimension = New Dimension() {Dimension.customVarName1, Dimension.customVarValue1}
            Dim objMetrics() As Metric = New Metric() {Metric.visits, Metric.pageviews, Metric.pageviewsPerVisit, Metric.avgTimeOnSite, Metric.percentNewVisits, Metric.visitBounceRate, Metric.goal3Completions, Metric.goal16Completions}

            Dim dims As StringBuilder = New StringBuilder()
            For Each objItemD In objDimension
                dims.Append("ga:" + objItemD.ToString() + ",")
            Next

            Dim mets As StringBuilder = New StringBuilder()
            For Each objItemM In objMetrics
                mets.Append("ga:" + objItemM.ToString() + ",")
            Next

            Dim requestUrl As String

            requestUrl = String.Format(requestUrlVFormat, _
                                       "ga:" + profileID, _
                                       dims.ToString().Trim(",".ToCharArray()), _
                                       mets.ToString().Trim(",".ToCharArray()), _
                                       from.ToString("yyyy-MM-dd"), _
                                       todate.ToString("yyyy-MM-dd"), _
                                       dims.ToString().Trim(",".ToCharArray()), _
                                       sessionToken, offset)

            Return requestUrl
        End Function

        '=============================================================================================
        'SetClient(ByVal sEMail As String, ByVal sPassword As String) called by ReqestReportID()
        'Instantiate and populate properties for myService via NewService():
        'myService = New AskWebServices.ReportService 
        'Used to invoke web services from ASK API
        '=============================================================================================
        'Public Sub SetClient(ByVal sEMail As String, ByVal sPassword As String)
        '    Email = sEMail

        '    If Left(sPassword, 1) = "&" Then sPassword = ""
        '    If IsNothing(sPassword) Then sPassword = ""

        '    Password = idsDefault_Password
        '    If sPassword.Length > 0 Then Password = sPassword

        '    'Instantiate myService() = New ASKWebServices.ReportsService
        '    NewService()
        '    'Set timeout per documentation
        '    myService.Timeout = Integer.MaxValue
        '    myCMService.Timeout = Integer.MaxValue
        '    myAdService.Timeout = Integer.MaxValue

        'End Sub
        '=============================================================================================
        'BldFileName(ByVal client As Clients) called by ASKReport.ReqestReportID 
        'Return String with full name of file.
        'Check for existing downloaded file in the local directory structure via naming conventions
        'AD_TheMetropolitanMuseumofArt_2008_03_03.csv for a daily report or
        'AM_TheMetropolitanMuseumofArt_2008_03_03.csv for a report spanning multiple days
        '=============================================================================================
        Public Function BldFileName(ByVal Client As Clients) As String

            'Perform Directory work
            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(Client)
            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)
            'Perform File naming
            Dim sFileName As String = ""
            Dim MyFile As String = ""
            If Client.GA.MultiDayReport Then
                MyFile = "GAM_" & Replace(Client.Name, " ", "") & "_" & Client.GA.AccountID & Client.GA.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            Else
                MyFile = "GAD_" & Replace(Client.Name, " ", "") & "_" & Client.GA.AccountID & Client.GA.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            End If
            sFileName = myDir & MyFile

            Return sFileName
        End Function
        '=============================================================================================
        'BldFileNameV(ByVal client As Clients) called by GAReport.ReqestReportID 
        'Return String with full name of file for storing Visitor GA Data.
        'Check for existing downloaded file in the local directory structure via naming conventions
        'AD_TheMetropolitanMuseumofArt_2008_03_03.csv for a daily report or
        'AM_TheMetropolitanMuseumofArt_2008_03_03.csv for a report spanning multiple days
        '=============================================================================================
        Public Function BldFileNameV(ByVal Client As Clients) As String

            'Perform Directory work
            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(Client)
            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)
            'Perform File naming
            Dim sFileName As String = ""
            Dim MyFile As String = ""
            If Client.GA.MultiDayReport Then
                MyFile = "GAVM_" & Replace(Client.Name, " ", "") & "_" & Client.GA.AccountID & Client.GA.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            Else
                MyFile = "GAVD_" & Replace(Client.Name, " ", "") & "_" & Client.GA.AccountID & Client.GA.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            End If
            sFileName = myDir & MyFile

            Return sFileName
        End Function
        '=============================================================================================
        'DoesFileExist(ByVal client As Clients) called by frmMain.vb Start_ASK()
        'Return Boolean
        'Check for existing downloaded file in the local directory structure via naming conventions
        'AD_TheMetropolitanMuseumofArt_2008_03_03.csv for a daily report or
        'AM_TheMetropolitanMuseumofArt_2008_03_03.csv for a report spanning multiple days
        '=============================================================================================
        Public Function DoesFileExist(ByVal Client As Clients) As Boolean
            'Return False

            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(Client)
            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)

            Dim sFileName As String = ""
            Dim MyFile As String = ""
            If Client.GA.MultiDayReport Then
                MyFile = "GAM_" & Replace(Client.Name, " ", "") & Client.GA.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            Else
                MyFile = "GAD_" & Replace(Client.Name, " ", "") & Client.GA.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            End If
            sFileName = myDir & MyFile

            Return File.Exists(sFileName)
        End Function


        '=============================================================================================
        'SaveFile(ByVal sUrl As String, ByVal sFileSave As String) called by DownloadReport()
        'Return Boolean
        'Using System.Net.WebClient DownloadFile(sUrl, sFileSave) method
        '=============================================================================================
        Public Function SaveFile(ByVal sUrl As String, ByVal sFileSave As String) As Boolean
            If File.Exists(sFileSave) Then File.Delete(sFileSave)

            Dim client As New WebClient
            client.DownloadFile(sUrl, sFileSave)

            Return File.Exists(sFileSave)
        End Function


        '======================================================================================
        ' GetMyDir - Build directory for Client                                                
        '======================================================================================
        Private Function GetMyDir(ByVal Client As Clients) As String

            Dim sDir As String
            If DirectoryNamesAsRecID Then
                sDir = SavePath & Client.CustID & "\"
            Else
                sDir = SavePath & Client.Name & "\"
            End If
            Return sDir
        End Function
        '======================================================================================
        ' ChkMoveFile - Move RAW JSON file into processed folder                                                       
        '======================================================================================
        Private Function ChkMoveFile(ByVal Client As Clients, ByVal sFile As String) As String
            Dim sDir As String = GetMyDir(Client)
            Dim oDir As New DirectoryInfo(sDir)
            Dim oFile As FileInfo

            If Directory.Exists(sDir) = False Then Directory.CreateDirectory(sDir)
            For Each oFile In oDir.GetFiles
                If sFile = oFile.FullName Then
                    MoveFile(oFile)
                End If

            Next
            Return True
        End Function
        Private Sub MoveFile(ByVal oFile As FileInfo)
            Dim sDir As String = oFile.DirectoryName & "\Processed\"
            Dim sFile As String = sDir & oFile.Name

            If Directory.Exists(sDir) = False Then Directory.CreateDirectory(sDir)
            If File.Exists(sFile) Then File.Delete(sFile)
            oFile.MoveTo(sFile)
        End Sub
#End Region

    End Class
End Namespace