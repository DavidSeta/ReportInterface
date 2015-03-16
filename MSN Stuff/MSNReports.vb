Imports System.IO
Imports System.Data
Imports ICSharpCode.SharpZipLib.Zip
Imports System.Net
Imports System.ServiceModel
Imports System.Runtime.Serialization
Imports System.Collections.Generic
Imports ReportInterface.MSNWebServices


Namespace MSNWebServices
    Public Class MSNReports
        Private myService As MSNWebServices.ReportingServiceClient
        Private queueRequest As MSNWebServices.SubmitGenerateReportRequest
        Private queueResponse As MSNWebServices.SubmitGenerateReportResponse
        Private statusResponse As PollGenerateReportResponse

        Private _BaseFilePath As String = ""
        Public DirectoryNamesAsRecID As Boolean = False

        'Private Const idsDefault_UserName As String = "API_MoreVisibility"
        'Private Const idsDefault_Password As String = "moreviz44"
        'Private Const idsDefault_UserAccessKey As String = "4030O59H"
        'OAuth 2.0 Testing Start
        'Private Const ClientID As String = "000000004011A289"
        'Private Const _tokenRefresh As String = "Clr53uA*659ytw2oBGFdgO9TseszWIbYysfIEKwu7dKqzmZhF4x5dELD48w7uHH2rbbhMMKLCLM2nYxmv6dEmGonRmQ4AR*MFgE7PbUAs1*v48o0qT2HZtD2x6GLnqVc8z3Qm7e5o7RENYftnHh3UQw20cqc1*Ki4LV3GxCoHjd0q08LnTgPuiw8Z7DxjOau*pYtS!2dGRYZiLk5VCnDmQ9CZZqXM70cmPSoV0ZlIZNHX0iuA8xW9cNOrA4RU6JuAdRz2m5PPyfN!y0H6SLnUtkrZdXGuRLFZbsHRLTb2UEVmJvjLulnAS1Cml3MreLSlqX4hznJaBBv9fCGYmg8CWFX!!ZGEVWTgIRpKcrngkNi"
        Private Const DesktopUri As String = "https://login.live.com/oauth20_desktop.srf"
        Private Const TokenUri As String = "https://login.live.com/oauth20_token.srf"
        Private Const RefreshUriFormatter As String = "{0}?client_id={1}&grant_type=refresh_token&redirect_uri={2}&refresh_token={3}"

        'OAuth 2.0 Testing End
        Public Username As String = ""
        Public Password As String = ""
        Public UserAccessKey As String = ""
        Public accessToken As String = ""
        Private appToken As String = ""
        Private devToken As String = ""
        Public _tokenRefresh As String = ""
        Public _clientID As String = ""


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
                Return BaseFilePath & "MSN\"
            End Get
        End Property

        Public Function DoesFileExist(ByVal Client As Clients) As Boolean

            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(Client)
            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)

            Dim sFileName As String = ""
            Dim MyFile As String = ""
            If Client.MSN.MultiDayReport Then
                myFile = "MM_" & Replace(Client.Name, " ", "") & Client.MSN.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            Else
                myFile = "MD_" & Replace(Client.Name, " ", "") & Client.MSN.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            End If
            sFileName = myDir & MyFile

            Return File.Exists(sFileName)
        End Function

#Region " Sub New (Overloaded)"
        Sub New(ByVal EngAuth As ClEngineSecurity)
            'V5 Stuff
            'App token reserved for future use

            'V6 now a string and already initialized
            'appToken = New ApplicationToken

            'V6 Credential all strings
            devToken = EngAuth.MSNAccessKey
            'OAuth2.0
            _tokenRefresh = EngAuth.MSNRefreshToken
            _clientID = EngAuth.MSNOAuthID

            'V6 Credentials back to strings
            'userCreds = New UserCredentials
            'userCreds.Username = EngAuth.MSNUserName
            'userCreds.Password = EngAuth.MSNPassword
            'OAuth Start
            Dim uriR = String.Format(RefreshUriFormatter, TokenUri, _clientID, DesktopUri, _tokenRefresh)
            If accessToken = "" Then accessToken = GetAccessTokens(uriR)
            'OAuth End
            'V4 stuff and not V6 stuff
            If Username = "" Then Username = EngAuth.MSNUserName
            If Password = "" Then Password = EngAuth.MSNPassword
            If UserAccessKey = "" Then UserAccessKey = EngAuth.MSNAccessKey


            NewService()
        End Sub
        Public Function GetAccessTokens(uri As String) As String
            Dim response As HttpWebResponse
            Dim request As HttpWebRequest
            Dim strJSON As String

            Dim sr As StreamReader
            Dim tokenResponse As String = Nothing
            'Convert String URI to a Uri
            Try
                Dim realUri = New Uri(uri, UriKind.Absolute)
                Dim addy = realUri.AbsoluteUri.Substring(0, realUri.AbsoluteUri.Length - realUri.Query.Length)
                request = HttpWebRequest.Create(addy)

                request.Method = "POST"
                request.ContentType = "application/x-www-form-urlencoded"
                request.GetRequestStream()
                Using writer = New StreamWriter(request.GetRequestStream())
                    writer.Write(realUri.Query.Substring(1))
                End Using

                response = request.GetResponse()

                ' The response is JSON, for example: {
                '                                     "token_type":"bearer",
                '                                     "expires_in":3600,
                '                                     "scope":"bingads.manage",
                '                                     "access_token":"<AccessToken>",
                '                                     "refresh_token":"<RefreshToken>"
                '                                    }

                ' Use the JSON serializer to serialize the response into the AccessTokens object.
                sr = New StreamReader(response.GetResponseStream())
                strJSON = sr.ReadToEnd()
                Dim values As Dictionary(Of String, String) = ParseJsonObjectResponse(strJSON)
                If values.ContainsKey("access_token") Then
                    tokenResponse = values("access_token")
                End If

            Catch e As Exception
                Log("MSN OAuth ERROR ON " & e.Message)
                'Dim response = DirectCast(e.Response, HttpWebResponse)
                'Console.WriteLine("HTTP status code: " & Convert.ToString(response.StatusCode))
            End Try

            Return tokenResponse
        End Function
        '==================================================================================================================================
        'ParseJsonObjectResponse() - New to V201309 - Request parsing                                                                               
        '                  Returns Hashtable - with value pairs from request                                                              
        '=================================================================================================================================
        ''' <summary>
        ''' Parses a json object response.
        ''' </summary>
        ''' <param name="contents">The json contents.</param>
        ''' <returns>A dictionary of key-value pairs.</returns>
        ''' <remarks>This is not a full-blown json parser, it can handle only cases
        ''' where the response is a json object without nested objects or arrays.
        ''' </remarks>
        Public Function ParseJsonObjectResponse(contents As String) As Dictionary(Of String, String)
            Dim retval As New Dictionary(Of String, String)()
            Dim splits As String() = contents.Trim(New Char() {" "c, "{"c, "}"c}).Split(","c)
            For Each split As String In splits
                Dim subSplits As String() = split.Trim(New Char() {" "c, ControlChars.Cr, ControlChars.Lf}).Split(":"c)
                If subSplits.Length = 2 Then
                    retval.Add(subSplits(0).Trim(New Char() {""""c, " "c}), subSplits(1).Trim(New Char() {""""c, " "c}))
                End If
            Next
            Return retval
        End Function
        Sub New(ByVal sUserName As String, ByVal sPassword As String, ByVal sUserAccessKey As String)
            'V5 Stuff
            'App token reserved for future use
            'appToken = New ApplicationToken

            'devToken = New DeveloperToken
            devToken = sUserAccessKey

            'V6 Credentials back to strings
            'userCreds = New UserCredentials
            'userCreds.Username = sUserName
            'userCreds.Password = sPassword

            'V4 stuff and now V6
            Username = sUserName
            Password = sPassword
            'UserAccessKey = sUserAccessKey

            NewService()
        End Sub

        Private Sub NewService()
            'V5 New Methods
            'ReportingServiceClient will permit QueueReport method to send Request via QueueReportRequest
            myService = New MSNWebServices.ReportingServiceClient

            'QueueReportRequest object for sending via ReportingServiceClient obj above 
            'and the Req object built in CreateKeywordReportRequest() function below
            queueRequest = New MSNWebServices.SubmitGenerateReportRequest
            queueRequest.ApplicationToken = Nothing
            queueRequest.DeveloperToken = devToken
            'OAuth Start
            queueRequest.AuthenticationToken = accessToken
            queueRequest.CustomerAccountId = "27131"
            queueRequest.CustomerId = "29183"
            'queueRequest.UserName = Username
            'queueRequest.Password = Password
            'OAuth End

            'QueueReportResponse object for obtaining results from API methods invoked
            queueResponse = New MSNWebServices.SubmitGenerateReportResponse
            'GetReportStatusResponse
            statusResponse = New PollGenerateReportResponse

            'V4
            'myService = New MSNWebServices.Reporting

            'Dim AuthHeader As New MSNWebServices.ApiUserAuthHeader
            'AuthHeader.UserName = Username
            'AuthHeader.Password = Password
            'AuthHeader.UserAccessKey = UserAccessKey

            'myService.Credentials = System.Net.CredentialCache.DefaultCredentials
            'myService.ApiUserAuthHeaderValue = AuthHeader
             
        End Sub
#End Region 'Sub New  (Overloaded)

#Region " CreateKeywordReportRequest (Overloaded)"
        '================================================================================================================================
        'CreateKeywordReportRequest(ReportParameters)
        'ReportParamaters(param) define the AccountID, start/end dates
        'Returns a KeywordPerformanceReportRequest which will then be sent to the API in Function ReqestReportID(ByVal Client As Clients)
        '================================================================================================================================
        Public Function CreateKeywordReportRequest(ByVal param As ReportParameters) As MSNWebServices.KeywordPerformanceReportRequest
            Dim intActID() As Long = {0}

            intActID(0) = param.AccountID

            Dim Req As New MSNWebServices.KeywordPerformanceReportRequest
            'V5 Account handling
            Req.Scope = New AccountThroughAdGroupReportScope
            Req.Scope.AccountIds = intActID
            Req.ReportName = "KeywordReport " + param.AccountID.ToString

            'Req.Account = param.AccountID
            Req.Language = MSNWebServices.ReportLanguage.English
            Req.Format = param.OuputFormat

            'Date Range
            Dim time = New MSNWebServices.ReportTime()
            Dim dayStart As New MSNWebServices.Date()
            Dim dayEnd As New MSNWebServices.Date()
            Dim dtWorkDate As Date

            dtWorkDate = param.GetStartDate
            dayStart.Day = dtWorkDate.Day
            dayStart.Month = dtWorkDate.Month
            dayStart.Year = dtWorkDate.Year

            dtWorkDate = param.GetEndDate
            dayEnd.Day = dtWorkDate.Day
            dayEnd.Month = dtWorkDate.Month
            dayEnd.Year = dtWorkDate.Year

            time.CustomDateRangeStart = dayStart
            time.CustomDateRangeEnd = dayEnd
            Req.Time = time
            Req.Aggregation = ReportAggregation.Daily

            'Build report columns
            Array.Resize(Req.Columns, 14)
            Req.Columns(0) = KeywordPerformanceReportColumn.AccountName
            Req.Columns(1) = KeywordPerformanceReportColumn.CampaignName
            Req.Columns(2) = KeywordPerformanceReportColumn.AdGroupName
            Req.Columns(3) = KeywordPerformanceReportColumn.Keyword
            Req.Columns(4) = KeywordPerformanceReportColumn.TimePeriod
            Req.Columns(5) = KeywordPerformanceReportColumn.Impressions
            Req.Columns(6) = KeywordPerformanceReportColumn.Clicks
            Req.Columns(7) = KeywordPerformanceReportColumn.Ctr
            Req.Columns(8) = KeywordPerformanceReportColumn.AverageCpc
            Req.Columns(9) = KeywordPerformanceReportColumn.Spend
            Req.Columns(10) = KeywordPerformanceReportColumn.AveragePosition
            Req.Columns(11) = KeywordPerformanceReportColumn.ConversionRate
            Req.Columns(12) = KeywordPerformanceReportColumn.Conversions
            Req.Columns(13) = KeywordPerformanceReportColumn.CostPerConversion

            Return Req
        End Function

        Public Function CreateKeywordReportRequest(ByVal param As ReportParameters, ByVal nAccountID As Long) As MSNWebServices.KeywordPerformanceReportRequest
            param.AccountID = nAccountID
            Return Me.CreateKeywordReportRequest(param)
        End Function

        Public Function CreateKeywordReportRequest(ByVal nAccountID As Long) As MSNWebServices.KeywordPerformanceReportRequest
            Return Me.CreateKeywordReportRequest(New ReportParameters, nAccountID)
        End Function

        Public Function CreateKeywordReportRequest(ByVal nAccountID As Long, ByVal dSingleDate As Date) As MSNWebServices.KeywordPerformanceReportRequest
            Dim Param As New ReportParameters
            Param.StartDate = dSingleDate
            Param.EndDate = dSingleDate
            Param.AccountID = nAccountID
            Return Me.CreateKeywordReportRequest(Param)
        End Function

        Public Function CreateKeywordReportRequest(ByVal nAccountID As Long, ByVal dStartDate As Date, ByVal dEndDate As Date) As MSNWebServices.KeywordPerformanceReportRequest
            Dim Param As New ReportParameters
            Param.StartDate = dStartDate
            Param.EndDate = dEndDate
            Param.AccountID = nAccountID
            Return Me.CreateKeywordReportRequest(Param)
        End Function
#End Region 'CreateKeywordReportRequest (Overloaded)

        Private Function GetMyDir(ByVal Client As Clients) As String
            '----------------------------------------------------------------------------------------------------------------
            ' Changed 3/14/07 so directories are now RecID's instead of names
            '----------------------------------------------------------------------------------------------------------------
            Dim sDir As String
            If DirectoryNamesAsRecID Then
                sDir = SavePath & Client.CustID & "\"
            Else
                sDir = SavePath & Client.Name & "\"
            End If
            Return sDir
        End Function


        '===========================================================================================================================
        ' Invoked via frmMain.Start_MSN()
        ' Build our request for KeywordPerformanceReport in function 
        'CreateKeywordReportRequest(ByVal param As ReportParameters) As MSNWebServices.KeywordPerformanceReportRequest
        '===========================================================================================================================
        Public Function ReqestReportID(ByVal Client As Clients) As String
            'Return Variable
            Dim myID As String

            'Build the Request via CreateKeywordReportRequest function 
            Dim Request As MSNWebServices.KeywordPerformanceReportRequest = Me.CreateKeywordReportRequest(CLng(Client.MSN.AccountID), Client.MSN.StartDate, Client.MSN.EndDate)
            'V5 New
            'Build the ReportRequest to be sent to API. Instantiation of queueRequest populated the authentication data
            queueRequest.ReportRequest = Request

            ' Submit the report request. This will throw an exception if 
            ' an error occurs.
            queueResponse = myService.SubmitGenerateReport(queueRequest)
            myID = queueResponse.ReportRequestId

            'V4
            'Dim APIFlags As Integer = 0
            'Dim myID As String = myService.RequestBIReport(APIFlags, Request)

            Return myID
        End Function

        '===========================================================================================================================
        ' Invoked via frmMain.GetStatus_MSN()
        ' Build our request for GetReportStatusResponse
        '===========================================================================================================================
        Public Function GetReportStatus(ByVal Client As Clients) As String
            Dim APIFlags As Integer = 0
            Dim sResult As String = "UNKNOWN"

            ' Initialize the Timers for looping until the status is returned from our request
            Dim waitMinutes As Integer = 1
            Dim maxWaitMinutes As Integer = 2
            Dim startTime As DateTime = DateTime.Now
            Dim ElapsedMinutes As Integer = 0

            ' Initialize the GetReportStatusResponse object to Nothing in case 
            ' an error occurs. The error will be handled below.
            statusResponse = Nothing

            Try
                Do
                    'Get Report Status and URL to download report
                    Dim statusRequest As PollGenerateReportRequest = _
                        New PollGenerateReportRequest

                    statusRequest.ApplicationToken = Nothing
                    statusRequest.DeveloperToken = devToken
                    'OAuth Start
                    statusRequest.AuthenticationToken = accessToken
                    'statusRequest.UserName = Username
                    'statusRequest.Password = Password
                    'OAuth End
                    statusRequest.ReportRequestId = Client.MSN.ReportID

                    ' Get the status of the report.
                    statusResponse = myService.PollGenerateReport(statusRequest)

                    If (ReportRequestStatusType.Success = _
                        statusResponse.ReportRequestStatus.Status) Then
                        sResult = "READY"
                        'Save the URL of the report to download
                        Client.MSN.URL = statusResponse.ReportRequestStatus.ReportDownloadUrl

                        ' The report is ready - so get out now.
                        Exit Do
                    ElseIf (ReportRequestStatusType.Pending = _
                            statusResponse.ReportRequestStatus.Status) Then
                        sResult = "IN PROGRESS"
                        ' The report is not yet ready. Loop again.
                        ' Wait 5 seconds before polling again.
                        System.Threading.Thread.Sleep(5 * 1000)
                        ElapsedMinutes = DateTime.Now.Subtract(startTime).Minutes
                    Else
                        ' An error occurred.
                        Exit Do
                    End If
                Loop While (ElapsedMinutes < maxWaitMinutes)

                'V4 old logic
                'Dim ReportStatus As MSNWebServices.APIReportStatusType = myService.GetAPIReportStatus(APIFlags, Client.MSN.ReportID)
                'Dim Status As MSNWebServices.APIStatusType = ReportStatus.ReportStatus

                'Select Case Status
                '    Case MSNWebServices.APIStatusType.REQUEST_IN_PROGRESS : sResult = "IN PROGRESS"
                '    Case MSNWebServices.APIStatusType.REQUEST_PENDING : sResult = "PENDING"
                '    Case MSNWebServices.APIStatusType.SUCCESS : sResult = "READY"
                '    Case MSNWebServices.APIStatusType.SUCCESS_PARTIAL : sResult = "READY"
                '    Case Else : sResult = "FAILED"
                'End Select

            Catch ex As Exception
                Debug.Print("Ex: " & ex.Message)
                Debug.Print("Inner: " & ex.InnerException.Message)
                'sResult = "FAILED"
            End Try

            Return sResult
        End Function

        '===========================================================================================================================
        ' Invoked via frmMain.Download_MSN_Once()
        ' Build our request for the downloading of the just created KeywordPerformanceReport
        '===========================================================================================================================
        Public Function DownloadReport(ByVal Client As Clients) As Boolean
            Dim bResult As Boolean = False
            Dim sFileName As String = ""
            Dim APIFlags As Integer = 0
            'V4 
            'Dim ReportStatus As MSNWebServices.APIReportStatusType = myService.GetAPIReportStatus(APIFlags, Client.MSN.ReportID)
            'Dim Status As MSNWebServices.APIStatusType = ReportStatus.ReportStatus

            'V5

            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(Client)

            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)

            Dim myFile As String = ""
            If Client.MSN.MultiDayReport Then
                myFile = "MM_" & Replace(Client.Name, " ", "") & Client.MSN.EndDate.ToString("_yyyy_MM_dd") & ".zip"
            Else
                myFile = "MD_" & Replace(Client.Name, " ", "") & Client.MSN.EndDate.ToString("_yyyy_MM_dd") & ".zip"
            End If
            sFileName = myDir & myfile

            Try
                Dim sURL As String = Client.MSN.URL
                'Dim sURL As String = ReportStatus.ReportDownloadURL    V4 Logic
                Log(Client.LogName(True) & "URL: " & sURL)
                If sURL <> "" Then
                    bResult = Me.saveFile(sURL, sFileName)
                    If bResult Then
                        Log("File Saved")
                    Else
                        Log("--- FILE NOT SAVED ---")
                    End If
                    UnZipFile(Client.MSN.ReportID, myDir, myFile, True)
                Else
                    Log("--- FILE NOT SAVED (NO URL) ---")
                End If
            Catch ex As Exception
                Log("MSN ERROR on " & Client.LogName(True) & ex.Message)

            End Try

            Return bResult
        End Function
        Public Function saveFile(ByVal sURl As String, ByVal sFIleSave As String) As Boolean
            If File.Exists(sFIleSave) Then File.Delete(sFIleSave)

            Dim client As New WebClient
            client.DownloadFile(sURl, sFIleSave)

            Return File.Exists(sFIleSave)
        End Function
        Public Sub UnZipFile(ByVal sReportID As String, ByVal sDir As String, ByVal sFile As String, ByVal bDeleteZipFile As Boolean)
            Dim sZip As String = sDir & sFile

            Dim fastZip As FastZip = New FastZip()
            fastZip.ExtractZip(sZip, sDir, Nothing)
            If bDeleteZipFile Then
                File.Delete(sZip)
            End If

            Dim sOld As String = sDir & sReportID & ".csv"
            Dim sNew As String = sDir & Left(sFile, sFile.Length - 4) & ".csv"

            File.Copy(sOld, sNew, True)
            File.Delete(sOld)

        End Sub
        '===========================================================================================================================
        'Invoked via frmMain.Start_MSN()
        'Close ReportingServiceClient
        '===========================================================================================================================
        Public Sub CloseMSNService()
            myService.Close()
        End Sub
        '==================================================================================================================================================
        'V4 - Removed going to V5
        'Public Function GetDaily(ByVal sClientID As String, ByVal dDate As Date) As Boolean
        '    Dim myID As String = ""
        '    Dim sURL As String = ""
        '    Dim bFailed As Boolean = False
        '    Dim nClick As Integer = 0
        '    Dim APIFlags As Integer = 0

        '    Debug.Print("")
        '    Debug.Print("---------------------------------")
        '    Debug.Print("    MSN Web Services")
        '    Debug.Print("---------------------------------")

        '    'Set to New User
        '    Debug.Print("UserID: " & sClientID)

        '    'Request Report
        '    Dim Request As New MSNWebServices.KeywordPerformanceReportRequest
        '    Request.StartDate = dDate
        '    Request.EndDate = dDate
        '    Request.AccountId = CInt(sClientID)

        '    myID = myService.RequestBIReport(APIFlags, Request)

        '    Debug.Print("myID: " & myID)
        '    If myID = "" Then
        '        Debug.Print("ID Error --  Failed Report")
        '        Return False
        '    End If

        '    'Wait Until Done
        '    While True
        '        Dim ReportStatus As MSNWebServices.APIReportStatusType = myService.GetAPIReportStatus(APIFlags, myID)
        '        Dim Status As MSNWebServices.APIStatusType = ReportStatus.ReportStatus
        '        If Status = MSNWebServices.APIStatusType.SUCCESS Then
        '            Debug.Print("Job Completed")
        '            sURL = ReportStatus.ReportDownloadURL
        '            Debug.Print("myURL: " & sURL)
        '            Exit While
        '        End If
        '        If Status > 3 Then
        '            bFailed = True
        '            Debug.Print("Job Failed")
        '            Exit While
        '        End If
        '        nClick = nClick + 1
        '        Debug.Print("Sleeping(" & nClick.ToString("000") & ") - Status: " & Status.ToString)
        '        System.Threading.Thread.Sleep(3000)
        '    End While

        '    'Download
        '    If sURL <> "" Then
        '        Dim bresult As Boolean = Me.saveFile(sURL, "C:\TestTim\M_" & myID.ToString & ".zip")
        '        Debug.Print("Download Success: " & bresult.ToString)
        '    End If

        '    'Unzip Download
        '    UnZipFile(myID.ToString, "C:\TestTim\", "C:\TestTim\" & ".zip", True)


        '    'All Done
        '    Debug.Print("All Done")
        '    Debug.Print("---------------------------------")
        '    Debug.Print("")

        '    Return True
        'End Function


        'V4 stuff removed
        'Public Function RequestReport(ByVal nAPIFlags As Integer, ByVal req As MSNWebServices.ReportRequest) As String
        '    Return myService.RequestBIReport(nAPIFlags, req)
        'End Function

        'Public Function ReportStatus(ByVal nAPIFlags As Integer, ByVal sReportID As String) As MSNWebServices.APIReportStatusType
        '    Dim oResult As MSNWebServices.APIReportStatusType = myService.GetAPIReportStatus(nAPIFlags, sReportID)

        '    Return oResult
        'End Function

    End Class

    '#Region " OAuth Token "
    '    ' The grant flow returns more fields than captured in this sample.
    '    ' Additional fields are not relevant for calling Bing Ads APIs or refreshing the token.

    '    <DataContract()> _
    '    Class AccessTokens
    '        ' Indicates the duration in seconds until the access token will expire.
    '        <DataMember()> _
    '        Friend expires_in As Integer = 0

    '        ' When calling Bing Ads service operations, the access token is used as  
    '        ' the AuthenticationToken header element.
    '        <DataMember()> _
    '        Friend access_token As String = Nothing

    '        ' May be used to get a new access token with a fresh expiration duration.
    '        <DataMember()> _
    '        Friend refresh_token As String = Nothing

    '        Public ReadOnly Property AccessToken() As String
    '            Get
    '                Return access_token
    '            End Get
    '        End Property
    '        Public ReadOnly Property ExpiresIn() As Integer
    '            Get
    '                Return expires_in
    '            End Get
    '        End Property
    '        Public ReadOnly Property RefreshToken() As String
    '            Get
    '                Return refresh_token
    '            End Get
    '        End Property
    '    End Class
    '#End Region 'Token

#Region " ReportParameters Class "
    Public Class ReportParameters
        Private _OuputFormat As MSNWebServices.ReportFormat

        Private _DateRangeType As MSNWebServices.ReportTime
        Private _AggregationType As MSNWebServices.ReportAggregation
        Private _Language As MSNWebServices.ReportLanguage

        Private _AccountID As Long = 0
        Private _StartDate As Date = Today
        Private _EndDate As Date = Today

        Sub New()
            'Set the Default Values
            OuputFormat = MSNWebServices.ReportFormat.Csv
            'DateRangeType = MSNWebServices.ReportTime.ReportTime
            AggregationType = MSNWebServices.ReportAggregation.Daily
            Language = MSNWebServices.ReportLanguage.English

            AccountID = 0
            StartDate = Today
            EndDate = Today
        End Sub

        Public Property OuputFormat() As MSNWebServices.ReportFormat

            Get
                Return _OuputFormat
            End Get
            Set(ByVal value As MSNWebServices.ReportFormat)
                _OuputFormat = value
            End Set
        End Property

        Public Property DateRangeType() As MSNWebServices.ReportTime
            Get
                Return _DateRangeType
            End Get
            Set(ByVal value As MSNWebServices.ReportTime)
                _DateRangeType = value
            End Set
        End Property

        Public Property AggregationType() As MSNWebServices.ReportAggregation
            Get
                Return _AggregationType
            End Get
            Set(ByVal value As MSNWebServices.ReportAggregation)
                _AggregationType = value
            End Set
        End Property

        Public Property Language() As MSNWebServices.ReportLanguage
            Get
                Return _Language
            End Get
            Set(ByVal value As MSNWebServices.ReportLanguage)
                _Language = value
            End Set
        End Property

        Public Property AccountID() As Long
            Get
                Return _AccountID
            End Get
            Set(ByVal value As Long)
                _AccountID = value
            End Set
        End Property

        Public Property StartDate() As Date
            Get
                Return _StartDate
            End Get
            Set(ByVal value As Date)
                _StartDate = value
            End Set
        End Property

        Public Property EndDate() As Date
            Get
                Return _EndDate
            End Get
            Set(ByVal value As Date)
                _EndDate = value
            End Set
        End Property

        Public Function GetStartDate() As Date

            Dim myDate As Date
            myDate = StartDate

            Return myDate
        End Function
        Public Function GetEndDate() As Date
            Dim myDate As Date

            myDate = EndDate

            Return myDate
        End Function
    End Class
#End Region 'ReportParameters Class
End Namespace