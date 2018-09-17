Imports System.IO
Imports System.Data
Imports System.Xml
Imports System.Net
Imports System.Text
Imports System.Web
Imports System.Reflection
'JSON Parser Requirement
Imports System.Web.Script.Serialization


Namespace GoogleWebServices
    Public Class GoogleReports
        Private WithEvents myService As GoogleWebServices.ReportDefinitionService
        '
        ' Oauth 2 - Generation of Access Token by exchanging Refresh Token
        '
        Public ws_AuthClient As AuthClient

        'New V201109 For AdHoc Reporting
        ' 
        'Account type to be used with ClientLogin API.
        ' 
        Private Const ACCOUNT_TYPE As String = "GOOGLE"

        ' 
        'Service type to be used with ClientLogin API.
        ' 
        Private Const SERVICE As String = "adwords"
        '
        ' OAuth2 URI
        '
        Private Const authOAFormat As String = "https://accounts.google.com/o/oauth2/token"
        Private Const authOARefreshToken As String = "refresh_token={0}&client_id={1}&client_secret={2}&grant_type=refresh_token"
        'Private Const authOABodyFormat As String = "code=4%2FcbH7EDcOLJbPRPAIuk8q_jotqVbu.4rWYve_iTXIYshQV0ieZDArw3U-AdAI&redirect_uri=urn:ietf:wg:oauth:2.0:oob&scope=https%3A%2F%2Fadwords.google.com%2Fapi%2Fadwords%2F&client_secret=vd8y42vwITiIisRmFidk7_XE&grant_type=authorization_code&client_id=151428027265.apps.googleusercontent.com"
        Public ClientID As String
        Public ClientSecret As String
        Public RefreshToken As String
        Private Settings As New clSettings

        ' 
        'The source identifier string to be used with ClientLogin API.
        ' 
        Private Const SOURCE As String = "PBJ-AWAPI-DotNetLib-v201601"
        Private Const URL As String = "https://adwords.google.com/api/adwords/reportdownload/v201710"
        'Private Const URL As String = "https://adwords.google.com/api/adwords/reportdownload/v201109"
        Private Const clientLogInUrl As String = "https://www.google.com/accounts/ClientLogin"
        Private Const authUrlFormat As String = "accountType=GOOGLE&Email={0}&Passwd={1}&source=Seta-AWAPI-DotNetLib-v201710&service=adwords"

        Private Const reportDef As String = "<reportDefinition><selector><fields>Date</fields>" + _
                "<fields>Id</fields><fields>CampaignName</fields><fields>AdGroupName</fields>" + _
                "<fields>Criteria</fields><fields>CriteriaType</fields><fields>Cost</fields>" + _
                "<fields>Impressions</fields><fields>Clicks</fields><fields>Cost</fields>" + _
                "<fields>Conversions</fields><fields>ViewThroughConversions</fields><fields>CustomerDescriptiveName</fields>" + _
                "<fields>AveragePosition</fields><fields>Ctr</fields><fields>AverageCpc</fields>" + _
                "<fields>ConversionRate</fields><fields>CostPerConversion</fields>" + _
                "<fields>ConversionsManyPerClick</fields>" + _
                "<fields>ConversionValue</fields>" + _
                "<dateRange><min>{0}</min><max>{1}</max></dateRange></selector>" + _
                "<reportName>Custom CRITERIA_PERFORMANCE_REPORT</reportName>" + _
                "<reportType>CRITERIA_PERFORMANCE_REPORT</reportType>" + _
                "<dateRangeType>CUSTOM_DATE</dateRangeType><downloadFormat>XML</downloadFormat>" + _
                "</reportDefinition>"

        'AdGroup Performance Report Definition

        Private Const reportDefAG As String = "<reportDefinition><selector><fields>Date</fields>" + _
                "<fields>Id</fields><fields>CampaignName</fields><fields>AdGroupName</fields>" + _
                "<fields>Criteria</fields><fields>CriteriaType</fields><fields>Cost</fields>" + _
                "<fields>Impressions</fields><fields>Clicks</fields><fields>Cost</fields>" + _
                "<fields>Conversions</fields><fields>ViewThroughConversions</fields><fields>CustomerDescriptiveName</fields>" + _
                "<fields>AveragePosition</fields><fields>Ctr</fields><fields>AverageCpc</fields>" + _
                "<fields>ConversionRate</fields><fields>CostPerConversion</fields>" + _
                "<fields>ConversionsManyPerClick</fields>" + _
                "<fields>ConversionValue</fields>" + _
                "<dateRange><min>{0}</min><max>{1}</max></dateRange></selector>" + _
                "<reportName>Custom DESTINATION_URL_REPORT</reportName>" + _
                "<reportType>DESTINATION_URL_REPORT</reportType>" + _
                "<dateRangeType>CUSTOM_DATE</dateRangeType><downloadFormat>XML</downloadFormat>" + _
                "</reportDefinition>"

        'End New V201109

        Private _BaseFilePath As String = ""
        Public DirectoryNamesAsRecID As Boolean = False

        Private idsDefault_Password As String
        Public Useragent As String
        Public Email As String
        Public Password As String
        Public Token As String
        Public AppToken As String
        Public DevToken As String
        'New for V201003
        Public AuthorizeToken As String = ""

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
                Return BaseFilePath & "Google\"
            End Get
        End Property

#Region " Sub New "

        '========================================================================================
        'New(ByVal engineAuthentication As ClEngineSecurity)
        'Called by frmMain_Load()
        '========================================================================================
        Sub New(ByVal engineAuthentication As ClEngineSecurity)
            'OAuth Refresh Token Work
            ws_AuthClient = New AuthClient(engineAuthentication)
            Token = ws_AuthClient.GetAuthHeader()
            'AdWords API Developer Token
            DevToken = engineAuthentication.GoogleToken


            ''OAuth2 Work
            ''Get AdWords authentication data from settings.xml 
            'Settings = New clSettings
            'ClientID = Settings.GoogleAdWordsClientID
            'ClientSecret = Settings.GoogleAdWordsClientSecret
            'RefreshToken = Settings.GoogleAdWordsRefreshToken

            'If Email = "" Then Email = engineAuthentication.GoogleEmail
            'If Password = "" Then Password = engineAuthentication.GooglePassword
            'If Useragent = "" Then Useragent = engineAuthentication.GoogleUserAgent

            ''OAuth2 logic for Token - implemented 10/2012
            'If Token = "" Then
            '    Token = GetToken(ClientID, ClientSecret, RefreshToken)
            '    AuthorizeToken = Token
            'End If

            ''Old logic for using ClientLogin to get token
            ''If Token = "" Then
            ''    Token = GetToken(Email, Password) 'If Token = "" Then Token = engineAuthentication.GoogleToken
            ''    AuthorizeToken = Token
            ''End If

            'If AppToken = "" Then AppToken = engineAuthentication.GoogleAppToken
            'idsDefault_Password = engineAuthentication.GooglePassword
            ''Required for V201109 and all future versions used to populate the header value developerToken
            'If DevToken = "" Then DevToken = engineAuthentication.GoogleToken

            'NewService()
        End Sub

#End Region 'Sub New

#Region " My Functions invoked via frmMain Start_Google()"
        '==============================================================================================
        'ReqestReports(ByVal client As Clients) Called by frmMain.vb Start_Google()                    
        'Previous to V201109 one used ReqestReportID(ByVal client As Clients)                          
        'Return String with ReportID or verbiage reflecting an error occurred while requesting: "ERROR"
        'Perform initialization   
        'SHOPPING_PERFORMANCE_REPORT - 9/3/2015 CHANGED
        '==============================================================================================
        Public Function ReqestReports(ByVal client As Clients) As String
            'Check if token set to expire
            Token = ws_AuthClient.GetAuthHeader()
            'Directory work
            Dim sFileName As String
            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(client)

            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)

            'File naming conventions work
            If client.Google.MultiDayReport Then
                sFileName = myDir & "GM_" & Replace(client.Name, " ", "") & client.Google.EndDate.ToString("_yyyy_MM_dd") & ".xml"
            Else
                sFileName = myDir & "GD_" & Replace(client.Name, " ", "") & client.Google.EndDate.ToString("_yyyy_MM_dd") & ".xml"
            End If

            Dim ReportID As String = "0"
            Dim myID As Integer = 0
            Dim myVID As Integer = 0

            Dim webRequest As WebRequest = HttpWebRequest.Create(URL)
          
            'We can use either the Client Google Email or the Google Client Account ID
            If Not String.IsNullOrEmpty(client.Google.GGLClientID) Then
                'webRequest.Headers.Add("Authorization: GoogleLogin auth=" + Token)  ClientLogIn
                webRequest.Headers.Add("Authorization: Bearer " + Token)
                webRequest.Headers.Add("clientCustomerId: " + client.Google.GGLClientID)
            Else
                Dim clAuthorizeToken As String = ""
                'clAuthorizeToken = New AuthClient(client.Google.AccountID, client.Google.AccountPWD).GetToken() ClientLog
                'webRequest.Headers.Add("Authorization: GoogleLogin auth=" + clAuthorizeToken)
                webRequest.Headers.Add("Authorization: Bearer " + Token)
                webRequest.Headers.Add("clientEmail: " + client.Google.AccountID)
            End If

            webRequest.Headers.Add("developerToken: " + DevToken)
            '  webRequest.Headers.Add("returnMoneyInMicros: true")  Depricated V201409 11-25-14
            webRequest.Method = "POST"
            webRequest.ContentType = "application/x-www-form-urlencoded"

            'Configure Report Definition with correct dates passed into tool
            Dim sreportDef As String
            sreportDef = String.Format(reportDef, client.Google.StartDate.ToString("yyyyMMdd"), client.Google.EndDate.ToString("yyyyMMdd"))

            'Dim postParams As String = "__rdxml=" + sreportDef
            Dim postParams As String = "__rdquery=SELECT Date, CampaignName, AdGroupName, OfferId, " & _
            "Impressions, Clicks, Cost, AverageCpc, Ctr, ConversionValue, Conversions, " & _
            "ConversionRate, CostPerConversion  FROM SHOPPING_PERFORMANCE_REPORT  WHERE Impressions > 0 " & _
            "DURING {0},{1}&__fmt=XML"
            postParams = String.Format(postParams, client.Google.StartDate.ToString("yyyyMMdd"), client.Google.EndDate.ToString("yyyyMMdd"))
            Dim postBytes As Byte() = Encoding.UTF8.GetBytes(postParams)
            webRequest.ContentLength = postBytes.Length
            Dim strmReq As Stream = webRequest.GetRequestStream()
            strmReq.Write(postBytes, 0, postBytes.Length)
            strmReq.Close()
            Dim response As WebResponse = Nothing
            Try
                'Attempt to get XML Report 
                response = webRequest.GetResponse()
            Catch ex As Exception
                'Had problems with API - return an error indicator
                Log("GOOGLE ERROR ON " & client.LogName(True) & ex.Message)
                If myID = 0 And myVID = 0 Then Return "ERROR"
            End Try

            Dim responseStream As Stream = response.GetResponseStream()
            Dim reader As New StreamReader(responseStream)
            Dim sresponseString As String = reader.ReadToEnd()

            'Write to file
            Dim sw As StreamWriter
            sw = New StreamWriter(sFileName)
            sw.Write(sresponseString)
            sw.Dispose()

            Return "GOOD"
    
        End Function
       
        '=============================================================================================
        'DoesFileExist(ByVal client As Clients) called by frmMain.vb Start_Google()
        'Return Boolean
        'Check for existing downloaded file in the local directory structure via naming conventions
        'GD_TheMetropolitanMuseumofArt_2008_03_03.xml for a daily report or
        'GM_TheMetropolitanMuseumofArt_2008_03_03.xml for a report spanning multiple days
        '=============================================================================================
        Public Function DoesFileExist(ByVal Client As Clients) As Boolean

            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(Client)
            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)

            Dim sFileName As String = ""
            Dim MyFile As String = ""
            If Client.Google.MultiDayReport Then
                MyFile = "GM_" & Replace(Client.Name, " ", "") & Client.Google.EndDate.ToString("_yyyy_MM_dd") & ".xml"
            Else
                MyFile = "GD_" & Replace(Client.Name, " ", "") & Client.Google.EndDate.ToString("_yyyy_MM_dd") & ".xml"
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
            client.Headers.Add("returnMoneyInMicros: true")

            client.DownloadFile(sUrl, sFileSave)

            Return File.Exists(sFileSave)
        End Function
      
        '======================================================================================
        ' Changed 3/14/07 so directories are now RecID's instead of names
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

#End Region

#Region " Output Properties (ReadOnly)"
        'Public ReadOnly Property responseTimeValue() As String
        '    Get
        '        Return myService.responseTimeValue.Text(0)
        '    End Get
        'End Property

        'Public ReadOnly Property unitsValue() As String
        '    Get
        '        Return myService.unitsValue.Text(0)
        '    End Get
        'End Property

        'Public ReadOnly Property operationsValue() As String
        '    Get
        '        Return myService.operationsValue.Text(0)
        '    End Get
        'End Property

        'Public ReadOnly Property requestIdValue() As String
        '    Get
        '        Return myService.requestIdValue.Text(0)
        '    End Get
        'End Property
#End Region 'Output Properties (ReadOnly)

#Region " Service Funtions - myService"
        ''===========================================================================================
        ''RequestKeywordReport (StartDate, EndDate) Invoked by GoogleReports.vb RequestReportID()
        ''Returning ReportID
        ''Determine if single date or date range and use:
        ''1)RequestDailyKeywordReport() or 
        ''2)RequestMTDKeywordReport() to instantiate the GoogleWebServices.DefinedReportJob
        ''Define properties for report via Me.Create_DailyKeyworkReportJob().
        ''Validate and schedule job via: 
        ''myService.validateReportJob(DefinedReportJob)
        ''myService.scheduleReportJob(DefinedReportJob)
        ''===========================================================================================
        'Public Function RequestKeywordReport(ByVal StartDate As Date, ByVal EndDate As Date) As Long
        '    If StartDate = EndDate Then Return RequestDailyKeywordReport(StartDate, EndDate)
        '    Return RequestMTDKeywordReport(StartDate, EndDate)
        'End Function
        ''===========================================================================================
        ''RequestMTDKeywordReport(ByVal StartDate As Date, ByVal EndDate As Date)
        ''Return ReportID via RequestReport(job)
        ''Define job properties for report via Me.Create_DailyKeyworkReportJob()
        ''===========================================================================================
        'Private Function RequestMTDKeywordReport(ByVal StartDate As Date, ByVal EndDate As Date) As Long
        '    Dim job As GoogleWebServices.DefinedReportJob = Me.Create_DailyKeyworkReportJob(StartDate, EndDate, "MyMultiDateReport")
        '    Return RequestReport(job)
        'End Function
        ''===========================================================================================
        ''RequestDailyKeywordReport(ByVal StartDate As Date, ByVal EndDate As Date)
        ''Return ReportID via RequestReport(job)
        ''Define job properties for report via Me.Create_DailyKeyworkReportJob()
        ''===========================================================================================
        'Private Function RequestDailyKeywordReport(ByVal StartDate As Date, ByVal EndDate As Date) As Long
        '    Dim job As GoogleWebServices.DefinedReportJob = Me.Create_DailyKeyworkReportJob(StartDate, EndDate, "MyDailyReport")
        '    Return RequestReport(job)
        'End Function
        ''===========================================================================================
        ''RequestReport(job) invoked via RequestMTDKeywordReport() and RequestDailyKeywordReport()
        ''Return ReportID
        ''Validate and Request Report Job via API:
        ''1)myService.validateReportJob(job) 
        ''2)myService.scheduleReportJob(job)
        ''===========================================================================================
        'Public Function RequestReport(ByVal job As GoogleWebServices.DefinedReportJob) As Long
        '    Dim lReportRequestID As Long = 0

        '    'Serialization and file variables
        '    'Dim xmlSerial As Serialization.XmlSerializer
        '    'Dim stream As Stream
        '    'Serialize the report request parameter to a file.
        '    'xmlSerial = New Serialization.XmlSerializer(job.GetType())
        '    'stream = New FileStream("C:\mysoap\ReportRequest.xml", FileMode.Create, FileAccess.Write, FileShare.None)
        '    'xmlSerial.Serialize(stream, job)
        '    'stream.Close()
        '    'Your report request variable has now been serialized to 
        '    'to the file specified above.

        '    'If we have any errors, probably due to the conversion data being requested, so try to validate w/o conv data 
        '    Try
        '        myService.validateReportJob(job)
        '    Catch ex As Exception
        '        job.selectedColumns = New String() {"Campaign", "AdGroup", "Keyword", "KeywordTypeDisplay", "AdWordsType", _
        '                    "Impressions", "Clicks", "CTR", "CPC", "Cost", "AveragePosition"}
        '        Try
        '            myService.validateReportJob(job)
        '        Catch ex1 As Exception
        '            Log("GOOGLE ERROR ON validateReportJob " & ex1.Message)
        '            Return lReportRequestID
        '        End Try
        '    End Try

        '    'Generate(Report)
        '    lReportRequestID = myService.scheduleReportJob(job)

        '    Return lReportRequestID
        'End Function

        ''getReportJobStatus
        'Public Function GetJobStatus(ByVal JobID As Long) As GoogleWebServices.ReportJobStatus
        '    Return myService.getReportJobStatus(JobID)
        'End Function

        ''GetAllJobs
        'Public Function GetAllJobStatus() As GoogleWebServices.ReportJob()
        '    Return myService.getAllJobs()
        'End Function

        ''getReportDownloadUrl
        'Public Function GetURL(ByVal JobID As Long) As String
        '    Return myService.getReportDownloadUrl(JobID)
        'End Function

        ''GetGzipReportDownloadUrl
        'Public Function GetGzipURL(ByVal JobID As Long) As String
        '    Return myService.getGzipReportDownloadUrl(JobID)
        'End Function

        ''DeleteReport
        'Public Sub DeleteReport(ByVal JobID As Long)
        '    myService.deleteReport(JobID)
        'End Sub
#End Region 'Service Funtions

#Region " Old Logic"


        'Sub New(ByVal sUserAgent As String, ByVal sEmail As String, ByVal sPassword As String, ByVal sToken As String, ByVal sAppToken As String)
        '    Useragent = sUserAgent
        '    Email = sEmail
        '    Password = sPassword
        '    Token = sToken
        '    AppToken = sAppToken

        '    NewService()
        'End Sub

        'Sub New(ByVal sEmail As String, ByVal sPassword As String, ByVal engineAuthentication As ClEngineSecurity)
        '    If Useragent = "" Then Useragent = engineAuthentication.GoogleUserAgent
        '    If Token = "" Then
        '        Token = GetToken(Email, Password) 'If Token = "" Then Token = engineAuthentication.GoogleToken
        '        AuthorizeToken = Token
        '    End If

        '    If AppToken = "" Then AppToken = engineAuthentication.GoogleAppToken

        '    Email = sEmail
        '    Password = sPassword

        '    NewService()
        'End Sub
        'Sub NewService()
        '    ' New AuthToken introduced in V201003
        '    'Check if we have already recieved the authToken
        '    '
        '    'Report Definition Service
        '    'myService = New GoogleWebServices.ReportDefinitionService()
        '    'myService.RequestHeader = New GoogleWebServices.RequestHeader()
        '    'myService.RequestHeader.authToken = AuthorizeToken
        '    'myService.RequestHeader.userAgent = Useragent
        '    'myService.RequestHeader.developerToken = Token

        '    ''
        '    'If AuthorizeToken = "" Then
        '    '    AuthorizeToken = New AuthClient(Email, Password).GetToken()
        '    'End If

        '    'myService = New GoogleWebServices.ReportService
        '    ''
        '    '' Populate authToken in service
        '    ''
        '    'myService.useragentValue = New GoogleWebServices.useragent
        '    'myService.emailValue = New GoogleWebServices.email
        '    'myService.passwordValue = New GoogleWebServices.password
        '    'myService.developerTokenValue = New GoogleWebServices.developerToken
        '    'myService.applicationTokenValue = New GoogleWebServices.applicationToken

        'End Sub
        ''========================================================================================
        ''GetToken(ByVal email, ByVal password) Return AuthToken String
        ''Called by New()
        ''OAuth2 - logic deployed 10/2012
        ''========================================================================================
        'Public Function GetToken(ByVal ClientID As String, ByVal ClientSecret As String, ByVal RefreshToken As String) As String
        '    Dim strJSON As String
        '    Dim Tokens() As String
        '    Dim item As String
        '    Dim tbOutput As StringBuilder
        '    tbOutput = New StringBuilder()

        '    Dim indentLevel As Integer = 0
        '    'Once we complete initial authorization we use authOARefreshToken and not authOABodyFormat
        '    Dim authBody As String = String.Format(authOARefreshToken, RefreshToken, ClientID, ClientSecret)

        '    'Client Log-in POST
        '    'Dim authBody As String = String.Format(authUrlFormat, Email, Password)

        '    'Dim authBody As String = authOABodyFormat
        '    Dim req As HttpWebRequest
        '    Dim response As HttpWebResponse

        '    Dim ser As JavaScriptSerializer
        '    'Dim dict = New Dictionary(Of String, Object)()

        '    Dim stream As Stream
        '    Dim sw As StreamWriter
        '    Dim sr As StreamReader

        '    'OAuth2.0
        '    'End point for all Token work
        '    'Via Refresh Token, we get and Access Token and use for one hour.
        '    req = HttpWebRequest.Create("https://accounts.google.com/o/oauth2/token")
        '    req.Method = "POST"
        '    req.ContentType = "application/x-www-form-urlencoded"
        '    req.UserAgent = "GoogleAdWordsReportInterface"

        '    stream = req.GetRequestStream()
        '    sw = New StreamWriter(stream)
        '    sw.Write(authBody)
        '    sw.Close()
        '    sw.Dispose()

        '    response = req.GetResponse()
        '    sr = New StreamReader(response.GetResponseStream())
        '    strJSON = sr.ReadToEnd()

        '    'JSON Work for finding access_token
        '    ser = New JavaScriptSerializer

        '    Dim dict As Dictionary(Of String, Object) = ser.Deserialize(Of Dictionary(Of String, Object))(strJSON)

        '    For Each strKey As String In dict.Keys
        '        If strKey.StartsWith("access_token") Then
        '            Dim o As Object = dict(strKey)
        '            Return o.ToString()
        '        End If
        '    Next
        '    'Dim objDeserialized As Object = ser.DeserializeObject(strJSON)


        '    'dict = ser.Deserialize(Of Dictionary(Of String, Object))(Token)
        '    'Tokens = Token.Split(ControlChars.CrLf.ToCharArray)
        '    'For Each item In Tokens
        '    '    If item.StartsWith("Auth=") Then
        '    '        Return item.Replace("Auth=", "")
        '    '    End If
        '    'Next

        '    Return String.Empty

        'End Function
        '========================================================================================
        'GetToken(ByVal email, ByVal password) Return AuthToken String
        'Called by New()
        'This is ClientLogin Authentication - no longer used as of 10/2012 and OAuth2 implementation
        '========================================================================================
        'Public Function GetToken(ByVal email As String, ByVal password As String) As String
        '    Dim Token As String
        '    Dim Tokens() As String
        '    Dim item As String

        '    Dim authBody As String = String.Format(authUrlFormat, email, password)
        '    Dim req As HttpWebRequest
        '    Dim response As HttpWebResponse

        '    Dim stream As Stream
        '    Dim sw As StreamWriter
        '    Dim sr As StreamReader

        '    req = HttpWebRequest.Create(clientLogInUrl)
        '    req.Method = "POST"
        '    req.ContentType = "application/x-www-form-urlencoded"
        '    req.UserAgent = "GoogleAdwordsReportInterface"

        '    stream = req.GetRequestStream()
        '    sw = New StreamWriter(stream)
        '    sw.Write(authBody)
        '    sw.Close()
        '    sw.Dispose()

        '    response = req.GetResponse()
        '    sr = New StreamReader(response.GetResponseStream())
        '    Token = sr.ReadToEnd()
        '    Tokens = Token.Split(ControlChars.CrLf.ToCharArray)
        '    For Each item In Tokens
        '        If item.StartsWith("Auth=") Then
        '            Return item.Replace("Auth=", "")
        '        End If
        '    Next

        '    Return String.Empty

        'End Function
        ''===========================================================================================
        ''Create_DailyKeyworkReportJob (StartDate, EndDate, ReportName) 
        ''NewJob = New GoogleWebServices.DefinedReportJob will be returned.
        ''Define following properties:
        ''name - Reports Name "Keyword Report - Daily" via sReportName pararmeter
        ''endDay - dEndDate pararmeter  
        ''startDay - dStartDate pararmeter
        ''selectedReportType - "Keyword"
        ''aggregationType() - "Daily"  
        ''selectedColumns() - Columns to be returned
        ''adGroupStatuses - GoogleWebServices.AdGroupStatus() - Enabled, Paused & Deleted
        ''includeZeroImpression = False
        ''===========================================================================================
        'Public Function Create_DailyKeyworkReportJob(ByVal dStartDate As Date, ByVal dEndDate As Date, ByVal sReportName As String) As GoogleWebServices.DefinedReportJob
        '    Dim NewJob As New GoogleWebServices.DefinedReportJob

        '    'Dim Accounts As String() = {}
        '    'NewJob.clientEmails = Accounts
        '    'NewJob.crossClient = False
        '    'NewJob.crossClientSpecified = False

        '    'Dim CampaignIDs As Integer() = {}
        '    'NewJob.campaigns = CampaignIDs

        '    ' Name this report job
        '    NewJob.name = sReportName

        '    ' Set the start and end date for the report
        '    NewJob.endDay = dEndDate
        '    NewJob.startDay = dStartDate

        '    'New to V11
        '    NewJob.selectedReportType = "Keyword"
        '    NewJob.aggregationTypes = New String() {"Daily"}
        '    'NewJob.selectedColumns = New String() {"Campaign", "AdGroup", "Keyword", _
        '    '                            "Impressions", "Clicks", "CTR", "CPC", "Cost", _
        '    '                            "Conversions", "CostPerConverstion", "ConversionRate", _
        '    '                            "AveragePosition"}

        '    NewJob.selectedColumns = New String() {"Campaign", "AdGroup", "Keyword", "KeywordTypeDisplay", "AdWordsType", _
        '                                "Impressions", "Clicks", "CTR", "CPC", "Cost", "AveragePosition", _
        '                                "Conversions", "CostPerConverstion", "ConversionRate"}


        '    'V10 Logic Below
        '    'Dim myOptions As GoogleWebServices.CustomReportOption() = {GoogleWebServices.CustomReportOption.Campaign, _
        '    '                                                           GoogleWebServices.CustomReportOption.AdGroup, _
        '    '                                                           GoogleWebServices.CustomReportOption.Keyword, _
        '    '                                                           GoogleWebServices.CustomReportOption.Impressions, _
        '    '                                                           GoogleWebServices.CustomReportOption.Clicks, _
        '    '                                                           GoogleWebServices.CustomReportOption.Ctr, _
        '    '                                                           GoogleWebServices.CustomReportOption.Cpc, _
        '    '                                                           GoogleWebServices.CustomReportOption.Conversions, _
        '    '                                                           GoogleWebServices.CustomReportOption.CostPerConversion, _
        '    '                                                           GoogleWebServices.CustomReportOption.Cost, _
        '    '                                                           GoogleWebServices.CustomReportOption.AveragePosition, _
        '    '                                                           GoogleWebServices.CustomReportOption.ConversionRate}
        '    'NewJob.customOptions = myOptions

        '    Dim ags As GoogleWebServices.AdGroupStatus() = {GoogleWebServices.AdGroupStatus.Enabled, _
        '                                                    GoogleWebServices.AdGroupStatus.Paused, _
        '                                                    GoogleWebServices.AdGroupStatus.Deleted}
        '    NewJob.adGroupStatuses = ags

        '    '---------------------------------------------------------------------------
        '    ' In this case, use the default settings for a keyword report.
        '    ' The report will contain keywords that
        '    ' -- have any matching type
        '    ' -- can be shown in all situations (content pages and search results)
        '    ' -- have any status
        '    '---------------------------------------------------------------------------
        '    'Dim kws As KeywordStatus() = {KeywordStatus.Normal, _
        '    '                              KeywordStatus.OnHold, _
        '    '                              KeywordStatus.InTrial, _
        '    '                              KeywordStatus.Disapproved, _
        '    '                              KeywordStatus.Disabled, _
        '    '                              KeywordStatus.Deleted}
        '    'NewJob.keywordStatuses = kws
        '    '--------------------------------------------------------------------------
        '    NewJob.includeZeroImpression = False
        '    'NewJob.includeZeroImpressionSpecified = True   Removed for V11

        '    ' Set the aggregation period to Daily V10
        '    'NewJob.aggregationType = GoogleWebServices.AggregationType.Daily
        '    'NewJob.aggregationTypeSpecified = True


        '    Return NewJob
        'End Function
        ''=============================================================================================
        ''DownloadReport(ByVal client As Clients) called by frmMain.vb Download_Google_Once()
        ''Return Boolean
        ''=============================================================================================
        'Public Function DownloadReport(ByVal Client As Clients) As Boolean
        '    SetClient(Client.Google.AccountID, Client.Google.AccountPWD)
        '    Dim bResult As Boolean = False
        '    Dim sFileName As String = ""

        '    'Get URL
        '    Dim myID As Long = CLng(Client.Google.ReportID)

        '    'Directory work
        '    If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
        '    Dim myDir As String = Me.GetMyDir(Client)

        '    If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)

        '    'File naming conventions work
        '    If Client.Google.MultiDayReport Then
        '        sFileName = myDir & "GM_" & Replace(Client.Name, " ", "") & Client.Google.EndDate.ToString("_yyyy_MM_dd") & ".xml"
        '    Else
        '        sFileName = myDir & "GD_" & Replace(Client.Name, " ", "") & Client.Google.EndDate.ToString("_yyyy_MM_dd") & ".xml"
        '    End If

        '    'Removed for V201109 Development
        '    ''Perform file download from URL returned via ID; and save in directory/file name, sFileName, just built
        '    'Try
        '    '    Dim sURL = myService.getReportDownloadUrl(myID)
        '    '    Log(Client.LogName(True) & "URL: " & sURL)
        '    '    If sURL <> "" Then
        '    '        bResult = Me.SaveFile(sURL, sFileName)
        '    '        If bResult Then
        '    '            Log("File Saved")
        '    '        Else
        '    '            Log("--- FILE NOT SAVED ---")
        '    '        End If
        '    '    Else
        '    '        Log("--- FILE NOT SAVED (NO URL) ---")
        '    '    End If

        '    '    'Clear Report on WebService
        '    '    If myID <> 0 Then myService.deleteReport(myID)

        '    'Catch ex As Exception
        '    '    Log("GOOGLE DOWNLOAD ERROR on " & Client.LogName(True) & ex.Message)

        '    'End Try

        '    Return bResult
        'End Function
        ''=============================================================================================
        ''ReqestReportID(ByVal client As Clients) called by frmMain.vb Start_Google()
        ''Return String with ReportID or verbiage reflecting an error occurred while requesting: "ERROR"
        ''Perform initialization 
        ''=============================================================================================
        'Public Function ReqestReportID(ByVal client As Clients) As String

        '    'Instantiate and populate properties for myService via NewService():
        '    SetClient(client.Google.AccountID, client.Google.AccountPWD)

        '    Dim myRptService As GoogleWebServices.ReportDefinitionService
        '    myRptService = New GoogleWebServices.ReportDefinitionService()
        '    myRptService.RequestHeader = New GoogleWebServices.RequestHeader()
        '    myRptService.RequestHeader.authToken = AuthorizeToken
        '    myRptService.RequestHeader.userAgent = Useragent
        '    myRptService.RequestHeader.developerToken = Token

        '    'One time shot to check available fields on report
        '    Dim reportDefinitionFields As ReportDefinitionField() = myRptService.getReportFields(ReportDefinitionReportType.CRITERIA_PERFORMANCE_REPORT, True)
        '    Console.WriteLine("Available fields for report:")
        '    For Each rptFld As GoogleWebServices.ReportDefinitionField In reportDefinitionFields
        '        Console.WriteLine(rptFld.fieldName + "(" + rptFld.fieldType + ")" + " Can Filter: " + rptFld.canFilter.ToString)
        '    Next

        '    ''Create Predicate to filter report to include only over 0 impressions
        '    'Dim impPredicate As New Predicate()
        '    'Dim sZero As Integer = 0
        '    'impPredicate.field = "Impressions"
        '    'impPredicate.[operator] = PredicateOperator.GREATER_THAN
        '    'impPredicate.operatorSpecified = True
        '    'impPredicate.values = New String() {sZero.ToString()}

        '    ''Create selector specifies the Fields for the report and the date range along with other selection criteria
        '    'Dim selector As GoogleWebServices.Selector
        '    'selector = New GoogleWebServices.Selector()

        '    'selector.fields = New String() {"AdGroupId", "AdGroupName", "CampaignName", "Date", "KeywordText", "KeywordMatchType", "Impressions", "Clicks", _
        '    '"Ctr", "AverageCpc", "Cost", "AveragePosition", "Conversions", "ConversionRate", "CostPerConversion"}
        '    'selector.predicates = New Predicate() {impPredicate}

        '    ''Create the report
        '    'Dim reportDefinition As GoogleWebServices.ReportDefinition
        '    'reportDefinition = New GoogleWebServices.ReportDefinition
        '    'reportDefinition.reportName = "Keyword Performance Report-API"

        '    'reportDefinition.downloadFormat = DownloadFormat.XML
        '    'reportDefinition.downloadFormatSpecified = True

        '    'reportDefinition.reportType = ReportDefinitionReportType.CRITERIA_PERFORMANCE_REPORT
        '    'reportDefinition.reportTypeSpecified = True

        '    ''Check if we want yesterday or custom date
        '    'If client.Google.IsYesterday Then
        '    '    reportDefinition.dateRangeType = ReportDefinitionDateRangeType.YESTERDAY
        '    'Else
        '    '    selector.dateRange = New GoogleWebServices.DateRange()
        '    '    selector.dateRange.min = Format(client.Google.StartDate, "yyyyMMdd")
        '    '    selector.dateRange.max = Format(client.Google.EndDate, "yyyyMMdd")
        '    '    reportDefinition.dateRangeType = ReportDefinitionDateRangeType.CUSTOM_DATE
        '    'End If
        '    'reportDefinition.dateRangeTypeSpecified = True

        '    'reportDefinition.selector = selector

        '    ''Create the operations
        '    'Dim operation As GoogleWebServices.ReportDefinitionOperation
        '    'operation = New GoogleWebServices.ReportDefinitionOperation

        '    'operation.operand = reportDefinition
        '    'operation.operator = [Operator].ADD
        '    'operation.operatorSpecified = True

        '    'Dim operations As ReportDefinitionOperation() = New ReportDefinitionOperation() {operation}

        '    'Dim result As GoogleWebServices.ReportDefinition() = myService.mutate(operations)
        '    Dim myID As Long = 0
        '    'myID = result(0).id

        '    'Build report, validate and request via myService.scheduleReportJob(job)
        '    'Dim myID As Long = 0
        '    'Try
        '    '    myID = Me.RequestKeywordReport(client.Google.StartDate, client.Google.EndDate)

        '    'Catch ex As Exception
        '    '    Log("GOOGLE ERROR ON " & client.LogName(True) & ex.Message)
        '    '    myID = 0

        '    ''End Try
        '    'If myID = 0 Then Return "ERROR"

        '    Return myID.ToString





        '    '====
        '    'Clear any Previous reports
        '    'ClearAllReports(client)

        '    ''Build report, validate and request via myService.scheduleReportJob(job)
        '    'Dim myID As Long = 0
        '    'Try
        '    '    myID = Me.RequestKeywordReport(client.Google.StartDate, client.Google.EndDate)

        '    'Catch ex As Exception
        '    '    Log("GOOGLE ERROR ON " & client.LogName(True) & ex.Message)
        '    '    myID = 0

        '    'End Try
        '    'If myID = 0 Then Return "ERROR"

        '    'Return myID.ToString
        'End Function
        ''=============================================================================================
        ''ClearAllReports(ByVal client As Clients) called by ReqestReportID()
        ''=============================================================================================
        'Private Sub ClearAllReports(ByVal client As Clients)

        '    'Removed for V201109 Development
        '    'Try
        '    '    Dim currentJobs() As GoogleWebServices.ReportJob = myService.getAllJobs
        '    '    If IsNothing(currentJobs) Then Exit Sub

        '    '    Dim nCount As Integer
        '    '    For nCount = 0 To currentJobs.GetUpperBound(0)
        '    '        If currentJobs(nCount).status = GoogleWebServices.ReportJobStatus.Completed Then
        '    '            myService.deleteReport(currentJobs(nCount).id)
        '    '        End If
        '    '    Next
        '    'Catch ex As Exception
        '    '    Log("-------------------------------------------------------------------------")
        '    '    Log(client.LogName & ": Clearing OLD Reports")
        '    '    Log("GOOGLE ERROR on " & client.LogName(True) & ex.Message)
        '    '    Log("Above Error does NOT affect the process")
        '    '    Log("-------------------------------------------------------------------------")
        '    'End Try

        'End Sub
        ''=============================================================================================
        ''GetReportStatus(ByVal client As Clients) called by frmMain.vb GetStatus_Google()
        ''=============================================================================================
        'Public Function GetReportStatus(ByVal Client As Clients) As String
        '    SetClient(Client.Google.AccountID, Client.Google.AccountPWD)
        '    Dim sResult As String = "UNKNOWN"
        '    'Removed for V201109 Development
        '    'Try
        '    '    Dim myid As Long = CLng(Client.Google.ReportID)
        '    '    Dim status As GoogleWebServices.ReportJobStatus = myService.getReportJobStatus(myid)

        '    '    Select Case status
        '    '        Case GoogleWebServices.ReportJobStatus.Completed : sResult = "READY"
        '    '        Case GoogleWebServices.ReportJobStatus.Failed : sResult = "FAILED"
        '    '        Case GoogleWebServices.ReportJobStatus.InProgress : sResult = "IN PROGRESS"
        '    '        Case GoogleWebServices.ReportJobStatus.Pending : sResult = "PENDING"
        '    '    End Select

        '    'Catch ex As Exception
        '    '    Log("GOOGLE STATUS ERROR ON " & Client.LogName(True) & ex.Message)
        '    '    sResult = "FAILED"

        '    'End Try

        '    Return sResult
        'End Function

        ''=============================================================================================
        ''SetClient(ByVal sEMail As String, ByVal sPassword As String) called by ReqestReportID()
        ''Instantiate and populate properties for myService via NewService():
        ''myService = New GoogleWebServices.ReportService 
        ''Used to invoke web services from Google API
        ''=============================================================================================
        'Public Sub SetClient(ByVal sEMail As String, ByVal sPassword As String)
        '    Email = sEMail

        '    If Left(sPassword, 1) = "&" Then sPassword = ""
        '    If IsNothing(sPassword) Then sPassword = ""

        '    Password = idsDefault_Password
        '    If sPassword.Length > 0 Then Password = sPassword

        '    NewService()

        '    'Used in V13 and removed for V201109
        '    'myService.useragentValue.Text = New String() {Useragent}
        '    'myService.emailValue.Text = New String() {Email}
        '    'myService.passwordValue.Text = New String() {Password}
        '    'myService.developerTokenValue.Text = New String() {Token}
        '    'myService.applicationTokenValue.Text = New String() {AppToken}
        'End Sub
#End Region ' End of old logic 

    End Class
End Namespace
