Imports System.IO
Imports System.Data
Imports System.Xml
Imports System.Net
Imports System.Web

Namespace GoogleWebServices
    Public Class GoogleReports
        Private WithEvents myService As GoogleWebServices.ReportDefinitionService
        Private WithEvents myAdService As GoogleAdGroupWebServices.AdGroupService

        Private _BaseFilePath As String = ""
        Public DirectoryNamesAsRecID As Boolean = False

        Private idsDefault_Password As String
        Public Useragent As String
        Public Email As String
        Public Password As String
        Public Token As String
        Public AppToken As String
        'New for V201003
        Public _AuthorizeToken As String = ""
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
            If Email = "" Then Email = engineAuthentication.GoogleEmail
            If Password = "" Then Password = engineAuthentication.GooglePassword
            If Useragent = "" Then Useragent = engineAuthentication.GoogleUserAgent
            If Token = "" Then Token = engineAuthentication.GoogleToken
            If AppToken = "" Then AppToken = engineAuthentication.GoogleAppToken
            idsDefault_Password = engineAuthentication.GooglePassword

            NewService()
            'Check if we have already recieved the authToken
            If _AuthorizeToken = "" Then
               _AuthorizeToken = New AuthClient(Email, Password).GetToken()
            End If
        End Sub

        Sub NewService()
            'Report Definition Service
            myService = New GoogleWebServices.ReportDefinitionService()
            myService.RequestHeader = New GoogleWebServices.RequestHeader()
            myService.RequestHeader.authToken = _AuthorizeToken
            myService.RequestHeader.userAgent = Useragent
            myService.RequestHeader.developerToken = Token

            'AdGroup Service
            myAdService = New GoogleAdGroupWebServices.AdGroupService()
            myAdService.RequestHeader = New GoogleAdGroupWebServices.RequestHeader()
            myAdService.RequestHeader.authToken = _AuthorizeToken
            myAdService.RequestHeader.userAgent = Useragent
            myAdService.RequestHeader.developerToken = Token
        End Sub
        '==================================================================================================
        ' Old Signatures for 
        '==================================================================================================
        Sub New(ByVal sUserAgent As String, ByVal sEmail As String, ByVal sPassword As String, ByVal sToken As String, ByVal sAppToken As String)
            Useragent = sUserAgent
            Email = sEmail
            Password = sPassword
            Token = sToken
            AppToken = sAppToken

            NewService()
            'Check if we have already recieved the authToken
            If _AuthorizeToken = "" Then
                _AuthorizeToken = New AuthClient(Email, Password).GetToken()
            End If
        End Sub

        Sub New(ByVal sEmail As String, ByVal sPassword As String, ByVal engineAuthentication As ClEngineSecurity)
            If Useragent = "" Then Useragent = engineAuthentication.GoogleUserAgent
            If Token = "" Then Token = engineAuthentication.GoogleToken
            If AppToken = "" Then AppToken = engineAuthentication.GoogleAppToken

            Email = sEmail
            Password = sPassword

            NewService()
            'Check if we have already recieved the authToken
            If _AuthorizeToken = "" Then
                _AuthorizeToken = New AuthClient(Email, Password).GetToken()
            End If
        End Sub
#End Region 'Sub New

#Region " Output Properties (ReadOnly)"
        ''TESTING V2010
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
        ''END TESTING V2010
#End Region 'Output Properties (ReadOnly)

#Region " Service Funtions - myService"
        ''TESTING V2010
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
        ''END TESTING V2010
#End Region 'Service Funtions

#Region " My Functions invoked via frmMain Start_Google()"
        '=============================================================================================
        'ReqestReportID(ByVal client As Clients) called by frmMain.vb Start_Google()
        'Return String with ReportID or verbiage reflecting an error occurred while requesting: "ERROR"
        'Perform initialization 
        '=============================================================================================
        Public Function ReqestReportID(ByVal client As Clients) As String
            Dim myID As Long
            'Instantiate and populate properties for myService via NewService():
            SetClient(client.Google.AccountID, client.Google.AccountPWD)

            'Clear any Previous reports
            'Aug 2010 Development
            'ClearAllReports(client)

            Dim myRptService As GoogleWebServices.ReportDefinitionService
            myRptService = New GoogleWebServices.ReportDefinitionService()
            myRptService.RequestHeader = New GoogleWebServices.RequestHeader()
            myRptService.RequestHeader.authToken = _AuthorizeToken
            myRptService.RequestHeader.userAgent = Useragent
            myRptService.RequestHeader.developerToken = Token
            'If client.CustID = 851 Then
            '    myRptService.RequestHeader.authToken = New AuthClient(client.Google.AccountID, client.Google.AccountPWD).GetToken()
            'End If

            ''One time shot to check available fields on report
            'Dim reportDefinitionFields As ReportDefinitionField() = myRptService.getReportFields(ReportDefinitionReportType.KEYWORDS_PERFORMANCE_REPORT, True)
            'Console.WriteLine("Available fields for report:")
            'For Each rptFld As GoogleWebServices.ReportDefinitionField In reportDefinitionFields
            '    Console.WriteLine(rptFld.fieldName + "(" + rptFld.fieldType + ")" + " Can Filter: " + rptFld.canFilter.ToString)
            'Next

            'Create Predicate to filter report to include only over 0 impressions
            Dim impPredicate As New Predicate()
            Dim sZero As Integer = 0
            impPredicate.field = "Impressions"
            impPredicate.[operator] = PredicateOperator.GREATER_THAN
            impPredicate.operatorSpecified = True
            impPredicate.values = New String() {sZero.ToString()}

            'Create selector specifies the Fields for the report and the date range along with other selection criteria
            Dim selector As GoogleWebServices.Selector
            selector = New GoogleWebServices.Selector()
           
            selector.fields = New String() {"AdGroupId", "AdGroupName", "CampaignName", "Date", "KeywordText", "KeywordMatchType", "Impressions", "Clicks", _
            "Ctr", "AverageCpc", "Cost", "AveragePosition", "Conversions", "ConversionRate", "CostPerConversion", "ViewThroughConversions"}
            selector.predicates = New Predicate() {impPredicate}

            'Create the report
            Dim reportDefinition As GoogleWebServices.ReportDefinition
            reportDefinition = New GoogleWebServices.ReportDefinition
            reportDefinition.reportName = "Keyword Performance Report-API"

            reportDefinition.downloadFormat = DownloadFormat.XML
            reportDefinition.downloadFormatSpecified = True

            reportDefinition.reportType = ReportDefinitionReportType.KEYWORDS_PERFORMANCE_REPORT
            reportDefinition.reportTypeSpecified = True

            'Check if we want yesterday or custom date
            If client.Google.IsYesterday Then
                reportDefinition.dateRangeType = ReportDefinitionDateRangeType.YESTERDAY
            ElseIf client.Google.GoogleEOM Then
                reportDefinition.dateRangeType = ReportDefinitionDateRangeType.LAST_MONTH
            Else
                selector.dateRange = New GoogleWebServices.DateRange()
                selector.dateRange.min = Format(client.Google.StartDate, "yyyyMMdd")
                selector.dateRange.max = Format(client.Google.EndDate, "yyyyMMdd")
                reportDefinition.dateRangeType = ReportDefinitionDateRangeType.CUSTOM_DATE
            End If
            reportDefinition.dateRangeTypeSpecified = True

            reportDefinition.selector = selector

            'Create the operations
            Dim operation As GoogleWebServices.ReportDefinitionOperation
            operation = New GoogleWebServices.ReportDefinitionOperation

            operation.operand = reportDefinition
            operation.operator = [Operator].ADD
            operation.operatorSpecified = True

            Dim operations As ReportDefinitionOperation() = New ReportDefinitionOperation() {operation}
            myID = 0
            Try
                Dim result As GoogleWebServices.ReportDefinition() = myService.mutate(operations)
                myID = result(0).id
            Catch ex As Exception
                Log("GOOGLE ERROR ON " & client.LogName(True) & ex.Message)

            End Try
            'Build report, validate and request via myService.scheduleReportJob(job)
            'Dim myID As Long = 0
            'Try
            '    myID = Me.RequestKeywordReport(client.Google.StartDate, client.Google.EndDate)

            'Catch ex As Exception
            '    Log("GOOGLE ERROR ON " & client.LogName(True) & ex.Message)
            '    myID = 0

            ''End Try

            If myID = 0 Then Return "ERROR"

            Return myID.ToString
        End Function
        '=============================================================================================
        'SetClient(ByVal sEMail As String, ByVal sPassword As String) called by ReqestReportID()
        'Instantiate and populate properties for myService via NewService():
        'myService = New GoogleWebServices.ReportService 
        'Used to invoke web services from Google API
        '=============================================================================================
        Public Sub SetClient(ByVal sEMail As String, ByVal sPassword As String)
            Email = sEMail

            If Left(sPassword, 1) = "&" Then sPassword = ""
            If IsNothing(sPassword) Then sPassword = ""

            Password = idsDefault_Password
            If sPassword.Length > 0 Then Password = sPassword

            NewService()
            myService.RequestHeader.authToken = _AuthorizeToken
            myService.RequestHeader.userAgent = Useragent
            myService.RequestHeader.developerToken = Token
            myService.RequestHeader.clientEmail = sEMail

            'AdGroup Service
            myAdService.RequestHeader.authToken = _AuthorizeToken
            myAdService.RequestHeader.userAgent = Useragent
            myAdService.RequestHeader.developerToken = Token
            myAdService.RequestHeader.clientEmail = sEMail

        End Sub
       
        '=============================================================================================
        'GetReportStatus(ByVal client As Clients) called by frmMain.vb GetStatus_Google()
        '=============================================================================================
        Public Function GetReportStatus(ByVal Client As Clients) As String
            'SetClient(Client.Google.AccountID, Client.Google.AccountPWD)
            Dim sResult As String = "READY"

            'TESTING V2010
            'Try
            '    Dim myid As Long = CLng(Client.Google.ReportID)
            '    Dim status As GoogleWebServices.ReportJobStatus = myService.getReportJobStatus(myid)

            '    Select Case status
            '        Case GoogleWebServices.ReportJobStatus.Completed : sResult = "READY"
            '        Case GoogleWebServices.ReportJobStatus.Failed : sResult = "FAILED"
            '        Case GoogleWebServices.ReportJobStatus.InProgress : sResult = "IN PROGRESS"
            '        Case GoogleWebServices.ReportJobStatus.Pending : sResult = "PENDING"
            '    End Select

            'Catch ex As Exception
            '    Log("GOOGLE STATUS ERROR ON " & Client.LogName(True) & ex.Message)
            '    sResult = "FAILED"

            'End Try
            'END TESTING V2010

            Return sResult
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
        'DownloadReport(ByVal client As Clients) called by frmMain.vb Download_Google_Once()          
        'Return Boolean 
        '=============================================================================================
        '=============================================================================================
        Public Function DownloadReport(ByVal Client As Clients) As Boolean
            
            SetClient(Client.Google.AccountID, Client.Google.AccountPWD)
            'Get ReportID for URL
            Dim myID As String = Client.Google.ReportID

            Dim bResult As Boolean = False
            Dim sFileName As String = ""
            Dim sURL As String = "https://adwords.google.com/api/adwords/reportdownload?__rd=" & myID

            'Directory work
            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(Client)
            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)

            'File naming conventions work
            If Client.Google.MultiDayReport Then
                sFileName = myDir & "GM_" & Replace(Client.Name, " ", "") & Client.Google.EndDate.ToString("_yyyy_MM_dd") & ".xml"
            Else
                sFileName = myDir & "GD_" & Replace(Client.Name, " ", "") & Client.Google.EndDate.ToString("_yyyy_MM_dd") & ".xml"
            End If

            'TESTING V2010
            'Perform file download from URL returned via ID; and save in directory/file name, sFileName, just built
            Try
                'V2010 url for report is built with report id.
                'Dim sURL = myService.getReportDownloadUrl(myID)
                Log(Client.LogName(True) & "URL: " & sURL)
                If sURL <> "" Then
                    bResult = Me.SaveFile(sURL, sFileName)
                    If bResult Then
                        Log("File Saved")
                    Else
                        Log("--- FILE NOT SAVED ---")
                    End If
                Else
                    Log("--- FILE NOT SAVED (NO URL) ---")
                End If

                'Clear Report on WebService V2010
                'If myID <> 0 Then myService.deleteReport(myID)

            Catch ex As Exception
                Log("GOOGLE DOWNLOAD ERROR on " & Client.LogName(True) & ex.Message)

            End Try
            'END TESTING V2010

                Return bResult
        End Function

        '=============================================================================================
        'SaveFile(ByVal sUrl As String, ByVal sFileSave As String) called by DownloadReport()
        'Return Boolean
        'Using System.Net.WebClient DownloadFile(sUrl, sFileSave) method
        '=============================================================================================
        Public Function SaveFile(ByVal sUrl As String, ByVal sFileSave As String) As Boolean
            If File.Exists(sFileSave) Then File.Delete(sFileSave)
            Dim FileDate As Date = GetFileDate(sFileSave)  'YYYY-MM-DD
            Dim sFileSaveTemp As String = sFileSave & "_temp"

            Dim client As New WebClient
            client.Headers.Add("Authorization: GoogleLogin auth=" & _AuthorizeToken)
            client.Headers.Add("clientEmail", myService.RequestHeader.clientEmail)
            client.Headers.Add("returnMoneyInMicros: true")
            client.DownloadFile(sUrl, sFileSave)
            ''client.DownloadFile(sUrl, sFileSaveTemp)

            ''V2010 Need to resolve AdGroupID and add attributes to the node report/table/row 
            'Dim oADGrpRsp As GoogleAdGroupWebServices.AdGroup
            'oADGrpRsp = New GoogleAdGroupWebServices.AdGroup()

            'Dim htCMDesc As Hashtable = New Hashtable()
            'Dim htAdtoCM As Hashtable = New Hashtable()
            'Dim htAdDesc As Hashtable = New Hashtable()
            'Dim nCMID As Long
            'Dim nAdID As Long
            'Dim sCMDesc As String = String.Empty
            'Dim sAdDesc As String = String.Empty

            'Dim xmlDoc As XmlDocument
            'Dim xmlRowNodes As XmlNodeList
            'Dim xmlNode As XmlNode
            'Dim xmlTemp As XmlNode
            'Dim xmlAdGroupID As XmlNode
            'Dim xmlAttr As XmlAttribute

            'xmlDoc = New XmlDocument
            'xmlDoc.Load(sFileSaveTemp)
            'xmlRowNodes = xmlDoc.SelectNodes("report/table/row")

            ''Process all rows
            'For Each xmlNode In xmlRowNodes
            '    'Add Date attribute to each node
            '    xmlAttr = xmlDoc.CreateAttribute("date")
            '    xmlAttr.InnerText = Format(FileDate, "yyy-MM-dd")
            '    xmlNode.Attributes.Append(xmlAttr)

            '    'Housekeeping for Campaign and AdGroup ID 
            '    nCMID = 0
            '    nAdID = 0

            '    'Get AdgroupID
            '    xmlAdGroupID = xmlNode.SelectSingleNode("@adGroupID")
            '    If Not xmlAdGroupID Is Nothing Then

            '        'Do we have the AdGroup Description already; if not get it in getAdGroup(nAdID)
            '        nAdID = Long.Parse(xmlAdGroupID.Value.ToString)
            '        If htAdDesc.ContainsKey(nAdID) Then
            '            sAdDesc = htAdDesc(nAdID).ToString
            '            xmlAttr = xmlDoc.CreateAttribute("adgroup")
            '            xmlAttr.InnerText = sAdDesc
            '            xmlNode.Attributes.Append(xmlAttr)
            '            'AdGroup to Campaign id mappings
            '            If htAdtoCM.ContainsKey(nAdID) Then
            '                nCMID = htAdtoCM(nAdID)
            '                sCMDesc = htCMDesc(nCMID).ToString
            '                xmlAttr = xmlDoc.CreateAttribute("campaign")
            '                xmlAttr.InnerText = sCMDesc
            '                xmlNode.Attributes.Append(xmlAttr)
            '            End If

            '        Else
            '            'Let's get the description for the AdGroupID, also Campaign Name

            '            Try
            '                oADGrpRsp = getAdGroup(nAdID)

            '                'Have names for campaign and adgroup
            '                'Place on HashTable for xref use 
            '                If Not oADGrpRsp Is Nothing Then

            '                    sAdDesc = oADGrpRsp.name.ToString
            '                    'Update hash table for AdGroup
            '                    htAdDesc.Add(nAdID, sAdDesc)

            '                    'Now add attribute for AdGroup
            '                    xmlAttr = xmlDoc.CreateAttribute("adgroup")
            '                    xmlAttr.InnerText = sAdDesc
            '                    xmlNode.Attributes.Append(xmlAttr)

            '                    'Campaign Information Returned
            '                    nCMID = oADGrpRsp.campaignId
            '                    sCMDesc = oADGrpRsp.campaignName

            '                    'Now add attribute for Campaign
            '                    xmlAttr = xmlDoc.CreateAttribute("campaign")
            '                    xmlAttr.InnerText = sCMDesc
            '                    xmlNode.Attributes.Append(xmlAttr)

            '                    'Now do Campaign lookup to add description
            '                    If htCMDesc.ContainsKey(nCMID) = False Then
            '                        htCMDesc.Add(nCMID, sCMDesc)
            '                    End If
            '                    'AdGroup to Campaign id xref mappings
            '                    If htAdtoCM.ContainsKey(nAdID) = False Then
            '                        htAdtoCM.Add(nAdID, nCMID)
            '                    End If

            '                Else
            '                    'If API to resolve AdGroupID fails, plug "unknown for Campaign and use ID in AdGroup
            '                    'Now add attribute for AdGroup
            '                    xmlAttr = xmlDoc.CreateAttribute("adgroup")
            '                    xmlAttr.InnerText = nAdID.ToString
            '                    xmlNode.Attributes.Append(xmlAttr)
            '                    'Now add attribute for Campaign
            '                    xmlAttr = xmlDoc.CreateAttribute("campaign")
            '                    xmlAttr.InnerText = "Unkown"
            '                    xmlNode.Attributes.Append(xmlAttr)
            '                End If
            '            Catch ex As Exception

            '                'If API to resolve AdGroupID fails, plug "unknown for Campaign and use ID in AdGroup
            '                Log("Google ERROR ON etting name for AdGroupID " & myService.RequestHeader.clientEmail & " " & nAdID.ToString & " " & ex.Message)
            '                xmlAttr = xmlDoc.CreateAttribute("adgroup")
            '                xmlAttr.InnerText = nAdID.ToString
            '                xmlNode.Attributes.Append(xmlAttr)
            '                'Now add attribute for Campaign
            '                xmlAttr = xmlDoc.CreateAttribute("campaign")
            '                xmlAttr.InnerText = "Unkown"
            '                xmlNode.Attributes.Append(xmlAttr)
            '            End Try

            '        End If

            '    End If  'xmlAdGroupID Not Null

            'Next 'For Each Looping

            ''Save temp file 
            'xmlDoc.Save(sFileSave)

            ''Need to move the temp file: sFileSaveTemp 
            'Dim oFile As New FileInfo(sFileSaveTemp)
            'MoveFile(oFile)
            ''V2010 END

            Return File.Exists(sFileSave)
        End Function

        '================================================================================================
        'getAdGroup(ByVal myAdGroupID As Long) called by GoogleReports.vb SaveFile() to resolve AdGroupID
        ' and get name for AdGroup and Campaign for report                                               
        '================================================================================================
        Public Function getAdGroup(ByVal myAdGroupID As Long) As GoogleAdGroupWebServices.AdGroup
            Dim myAdGrpService As GoogleAdGroupWebServices.AdGroupService
            myAdGrpService = New GoogleAdGroupWebServices.AdGroupService()
            myAdGrpService.RequestHeader = New GoogleAdGroupWebServices.RequestHeader()
            myAdGrpService.RequestHeader.authToken = _AuthorizeToken
            myAdGrpService.RequestHeader.userAgent = Useragent
            myAdGrpService.RequestHeader.developerToken = Token
            myAdGrpService.RequestHeader.clientEmail = myAdService.RequestHeader.clientEmail

            'Create selector specifies the Fields for the report and the date range along with other selection criteria
            Dim selector As GoogleWebServices.Selector
            selector = New GoogleWebServices.Selector()

            Dim Result As GoogleAdGroupWebServices.AdGroup
            Result = New GoogleAdGroupWebServices.AdGroup()
            Dim adgroup As GoogleAdGroupWebServices.AdGroup = Nothing
            adgroup = New GoogleAdGroupWebServices.AdGroup()

            Dim myArrayIDs() As Long = Nothing
            ReDim myArrayIDs(0)
            myArrayIDs(0) = myAdGroupID

            Dim ix As Long
            Try
                ' Get all AdGroups for specified campaign id passed.
                Dim adGroupSelector As GoogleAdGroupWebServices.AdGroupSelector
                adGroupSelector = New GoogleAdGroupWebServices.AdGroupSelector()
                adGroupSelector.adGroupIds = myArrayIDs

                Dim page As GoogleAdGroupWebServices.AdGroupPage = myAdGrpService.get(adGroupSelector)

                ' Display AdGroups.
                ix = 0
                If page IsNot Nothing AndAlso page.entries IsNot Nothing Then
                    If page.entries.Length > 0 Then
                        For Each adgroup In page.entries
                            Result = adgroup
                            'Console.WriteLine("AdGroups with id = '{0}', name = '{1}' and status = '{2}'" + " was found.", adgroup.id, adgroup.name, adgroup.status)

                        Next
                    Else
                        Console.WriteLine("No adgroups were found.")
                    End If
                End If
            Catch ex As Exception
                Console.WriteLine("Failed to retrieve adGroup(s). Exception says ""{0}""", ex.Message)
            End Try

            Return Result

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
        Public Function GetFileDate(ByVal sFileName As String) As Date
            Dim sTemp As String = sFileName.Split(".")(0)
            sTemp = Right(sTemp, 10)

            Dim myFileDate As Date = CDate(Left(sTemp, 4) & "-" & Mid(sTemp, 6, 2) & "-" & Right(sTemp, 2))

            Return myFileDate
        End Function
        Private Sub MoveFile(ByVal oFile As FileInfo)
            Dim sDir As String = oFile.DirectoryName & "\RawXML\"
            Dim sFile As String = sDir & oFile.Name

            If Directory.Exists(sDir) = False Then Directory.CreateDirectory(sDir)
            If File.Exists(sFile) Then File.Delete(sFile)
            oFile.MoveTo(sFile)
        End Sub

        'Aug 2010 Development
        ''=============================================================================================
        ''ClearAllReports(ByVal client As Clients) called by ReqestReportID()
        ''=============================================================================================
        'Private Sub ClearAllReports(ByVal client As Clients)

        '    Try
        '        Dim currentJobs() As GoogleWebServices.ReportJob = myService.getAllJobs
        '        If IsNothing(currentJobs) Then Exit Sub

        '        Dim nCount As Integer
        '        For nCount = 0 To currentJobs.GetUpperBound(0)
        '            If currentJobs(nCount).status = GoogleWebServices.ReportJobStatus.Completed Then
        '                myService.deleteReport(currentJobs(nCount).id)
        '            End If
        '        Next
        '    Catch ex As Exception
        '        Log("-------------------------------------------------------------------------")
        '        Log(client.LogName & ": Clearing OLD Reports")
        '        Log("GOOGLE ERROR on " & client.LogName(True) & ex.Message)
        '        Log("Above Error does NOT affect the process")
        '        Log("-------------------------------------------------------------------------")
        '    End Try

        'End Sub
        'TESTING V2010
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
        'END TESTING V2010

#End Region

    End Class
End Namespace
